import os
import re
import pandas as pd
from typing import List, Dict
from vertexai.generative_models import GenerativeModel, GenerationConfig, SafetySetting

class RuleExtractor:
    def __init__(self, model_name: str, temperature: float = 0.0, max_tokens: int = 8192):
        self.model = GenerativeModel(model_name)
        self.generation_config = GenerationConfig(
            temperature=temperature,
            max_output_tokens=max_tokens
        )
        self.safety_settings = [
            SafetySetting(
                category=SafetySetting.HarmCategory.HARM_CATEGORY_HATE_SPEECH,
                threshold=SafetySetting.HarmBlockThreshold.OFF
            ),
            SafetySetting(
                category=SafetySetting.HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT,
                threshold=SafetySetting.HarmBlockThreshold.OFF
            ),
            SafetySetting(
                category=SafetySetting.HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT,
                threshold=SafetySetting.HarmBlockThreshold.OFF
            ),
            SafetySetting(
                category=SafetySetting.HarmCategory.HARM_CATEGORY_HARASSMENT,
                threshold=SafetySetting.HarmBlockThreshold.OFF
            )
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
        try:
            match = re.search(r"\`\`\`(?:csv)?\n([\s\S]*?)\n\`\`\`", llm_output)
            if match:
                return match.group(1).strip()
            return ""
        except Exception:
            return ""

    def process_chunk(self, chunk: Dict, context_chunks: List[Dict]) -> pd.DataFrame:
        context_str = "\n".join([f"--- Name: {c['name']} ---\n{c['content']}" for c in context_chunks])
        prompt = self.prompt_template.format(
            context=context_str,
            name=chunk['name'],
            file_path=chunk['file_path'],
            block_type=chunk['block_type'],
            code=chunk['content']
        )

        try:
            response = self.model.generate_content(
                prompt,
                generation_config=self.generation_config,
                safety_settings=self.safety_settings
            ).text
            
            csv_str = self._extract_csv(response)
            if not csv_str:
                return pd.DataFrame()
            
            # Use StringIO directly to avoid deprecation warnings in newer pandas
            import io
            df = pd.read_csv(io.StringIO(csv_str), sep="|", header=0)
            # Add trace to original chunk id
            df['Source_Chunk_ID'] = chunk['id']
            return df
            
        except Exception as e:
            print(f"Error extracting rules for {chunk['name']}: {e}")
            return pd.DataFrame()
