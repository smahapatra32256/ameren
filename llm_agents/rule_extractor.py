import io
import os
import re
import time
from typing import Dict, List

import pandas as pd
from vertexai.generative_models import GenerationConfig, GenerativeModel, SafetySetting


RETRYABLE_ERROR_MARKERS = ("429", "503", "504", "500", "SSLError", "EOF", "timeout")
REQUIRED_COLUMNS = [
    "Sequence_Order",
    "Rule_Flow",
    "Rule_Task",
    "Rule_Name",
    "Rule",
    "Actual_Code",
    "Explanation",
]


class RuleExtractionError(Exception):
    def __init__(self, message: str, retryable: bool = False):
        super().__init__(message)
        self.retryable = retryable


class RuleExtractor:
    def __init__(self, model_name: str, temperature: float = 0.0, max_tokens: int = 8192):
        self.model = GenerativeModel(model_name)
        self.generation_config = GenerationConfig(
            temperature=temperature,
            max_output_tokens=max_tokens,
        )
        self.safety_settings = [
            SafetySetting(
                category=SafetySetting.HarmCategory.HARM_CATEGORY_HATE_SPEECH,
                threshold=SafetySetting.HarmBlockThreshold.OFF,
            ),
            SafetySetting(
                category=SafetySetting.HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT,
                threshold=SafetySetting.HarmBlockThreshold.OFF,
            ),
            SafetySetting(
                category=SafetySetting.HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT,
                threshold=SafetySetting.HarmBlockThreshold.OFF,
            ),
            SafetySetting(
                category=SafetySetting.HarmCategory.HARM_CATEGORY_HARASSMENT,
                threshold=SafetySetting.HarmBlockThreshold.OFF,
            ),
        ]
        self.prompt_template = """
You are an expert VB.NET developer and Business Analyst.
Your task is to extract business rules from the following VB.NET code block.
You will be provided with the target code chunk and potentially some related context code from FAISS.

<CONTEXT>
{context}
</CONTEXT>

<TARGET_CODE>
Block Name: {name}
File Path: {file_path}
Block Type: {block_type}

{code}
</TARGET_CODE>

Extract all underlying business rules, logic conditions, and operations found in TARGET_CODE sequentially.
If the code is technical (e.g. logging, UI bindings), note it as a technical rule.
Output MUST be strictly in CSV format enclosed in ```csv...``` with the following exact headers separated by "|":
Sequence_Order|Rule_Flow|Rule_Task|Rule_Name|Rule|Actual_Code|Explanation

Instructions:
- Sequence_Order: A chronological indicator (e.g., 1, 2, 3...) representing the step number of the rule within the flow.
- Actual_Code: The literal VB.NET code snippet (e.g. `amount <= 0`) that triggered the rule.

CRITICAL CSV FORMATTING RULES:
1. Keep the entire row on a SINGLE line. Do NOT use actual newlines (\n) inside `Actual_Code` or `Explanation`. Use a space instead of newlines.
2. If `Actual_Code` or `Explanation` contains the pipe character `|`, replace it with the word `OR` to prevent breaking the CSV delimiter.
3. Keep things simple without excessive double quotes.

Output Example:
```csv
Sequence_Order|Rule_Flow|Rule_Task|Rule_Name|Rule|Actual_Code|Explanation
1|{name}_Flow|Process_Data|Check_If_Null|If data is null, throw exception|If data Is Nothing Then|Ensures data integrity before processing
```
"""

    def _extract_csv(self, llm_output: str) -> str:
        match = re.search(r"```(?:csv)?\s*([\s\S]*?)\s*```", llm_output)
        if match:
            return match.group(1).strip()

        if "Sequence_Order|Rule_Flow|" in llm_output:
            return llm_output.strip()
        return ""

    def _normalize_method_name(self, name: str) -> str:
        # Backward compatibility for old chunk names (foo_part1)
        part_match = re.search(r"^(.*)_part\d+$", name)
        return part_match.group(1) if part_match else name

    def process_chunk(self, chunk: Dict, context_chunks: List[Dict]) -> pd.DataFrame:
        context_str = "\n".join([f"--- Name: {c['name']} ---\n{c['content']}" for c in context_chunks])
        prompt = self.prompt_template.format(
            context=context_str,
            name=chunk["name"],
            file_path=chunk["file_path"],
            block_type=chunk["block_type"],
            code=chunk["content"],
        )

        max_retries = 3
        last_error: str = "Unknown extraction failure"

        for attempt in range(max_retries):
            try:
                response_obj = self.model.generate_content(
                    prompt,
                    generation_config=self.generation_config,
                    safety_settings=self.safety_settings,
                )
                response_text = (response_obj.text or "").strip()
                csv_str = self._extract_csv(response_text)
                if not csv_str:
                    raise RuleExtractionError("LLM response did not contain a parseable CSV payload", retryable=False)

                df = pd.read_csv(
                    io.StringIO(csv_str),
                    sep="|",
                    header=0,
                    dtype=str,
                    keep_default_na=False,
                    on_bad_lines="skip",
                    engine="python",
                )
                if df.empty:
                    raise RuleExtractionError("Extracted CSV was empty after parsing", retryable=False)

                missing_columns = [c for c in REQUIRED_COLUMNS if c not in df.columns]
                if missing_columns:
                    raise RuleExtractionError(
                        f"Extracted CSV missing required columns: {', '.join(missing_columns)}",
                        retryable=False,
                    )

                # Keep schema strict and append trace metadata.
                df = df[REQUIRED_COLUMNS].copy()
                df["Source_File"] = os.path.basename(chunk["file_path"])
                df["Source_Method"] = self._normalize_method_name(chunk["name"])
                df["Source_Chunk_ID"] = chunk["id"]
                return df

            except RuleExtractionError as e:
                last_error = str(e)
                # Structural/format failures are non-retryable; fail fast for observability.
                raise
            except Exception as e:
                err_str = str(e)
                last_error = err_str
                is_retryable = any(code in err_str for code in RETRYABLE_ERROR_MARKERS)
                if is_retryable and attempt < max_retries - 1:
                    wait = 5 * (2 ** attempt)
                    time.sleep(wait)
                    continue
                raise RuleExtractionError(err_str, retryable=is_retryable)

        raise RuleExtractionError(last_error, retryable=False)
