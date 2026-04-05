import os
import re
import time
from typing import Dict, List

from vertexai.generative_models import GenerationConfig, GenerativeModel, SafetySetting


RETRYABLE_ERROR_MARKERS = ("429", "503", "504", "500", "SSLError", "EOF", "timeout")


class UMLGenerationError(Exception):
    def __init__(self, message: str, retryable: bool = False):
        super().__init__(message)
        self.retryable = retryable


class UMLGenerator:
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

        self.manifest_prompt = """
You are an expert VB.NET Architect.
Extract the execution flow of the following VB.NET code block into a PlantUML Sequence Diagram snippet.
We want to see exactly what this method does, what decisions it makes, and what external systems/databases it calls.

<CODE_BLOCK>
File Name: {file_name}
Block Name: {name}
Block Type: {block_type}

{code}
</CODE_BLOCK>

CRITICAL INSTRUCTION:
Return ONLY the raw PlantUML Sequence Diagram snippet for this specific method.
DO NOT wrap it in @startuml or @enduml.
DO NOT use single quotes `'` or comments.
DO NOT use the `create` keyword. Just use `A -> B: message` followed by `activate B`.
DO NOT use the `return` keyword. Explicitly use `B --> A: response` and `deactivate B`.
Use standard Sequence Diagram syntax.
Example format:
group {name}
  Caller -> "{file_name}": {name}(args)
  activate "{file_name}"
  alt if condition is true
    "{file_name}" -> Database: Execute Query
    activate Database
    Database --> "{file_name}": Result
    deactivate Database
  else
    "{file_name}" -> Logger: Log Error
  end
  deactivate "{file_name}"
end
Keep descriptions short and focus on the business action.
"""

        self.reduce_prompt = """
You are an expert VB.NET Architect.
Below is a collection of Sequence Diagram snippets extracted from methods in a single source file.
Your task is to REDUCE these snippets into a single, cohesive PlantUML Sequence Diagram for the entire file '{file_name}'.

<FILE_NAME>
{file_name}
</FILE_NAME>

<MANIFESTS>
{manifests}
</MANIFESTS>

CRITICAL RULES:
1. Generate exactly one `@startuml` ... `@enduml` block.
2. Declare the participants at the top (e.g., `participant "{file_name}"`, `database DB`).
3. Combine all the provided `group` blocks logically.
4. If methods call each other internally, ensure the sequence reflects that.
5. DO NOT use single quotes `'` or comments.
6. DO NOT use `activate Participant as Alias`. Only use `activate Participant` and `deactivate Participant`.
7. When using `alt`, ensure proper `else` and a single `end`. Do not nest `alt` without matching `end`.
8. DO NOT use the `create` or `return` keywords anywhere. Use explicit arrows (`->` and `-->`) and `activate`/`deactivate`.
9. DO NOT use Markdown wrappers like ```plantuml. Just the raw code.
"""

    def _generate_with_retries(self, prompt: str, config: GenerationConfig, max_retries: int = 3) -> str:
        last_error = "Unknown UML generation error"
        for attempt in range(max_retries):
            try:
                response_obj = self.model.generate_content(
                    prompt,
                    generation_config=config,
                    safety_settings=self.safety_settings,
                )
                text = (response_obj.text or "").strip()
                if not text:
                    raise UMLGenerationError("LLM returned an empty UML payload", retryable=False)
                return text
            except UMLGenerationError:
                raise
            except Exception as e:
                err_str = str(e)
                last_error = err_str
                is_retryable = any(code in err_str for code in RETRYABLE_ERROR_MARKERS)
                if is_retryable and attempt < max_retries - 1:
                    wait = 5 * (2 ** attempt)
                    time.sleep(wait)
                    continue
                raise UMLGenerationError(err_str, retryable=is_retryable)
        raise UMLGenerationError(last_error, retryable=False)

    def extract_manifest(self, chunk: Dict) -> str:
        prompt = self.manifest_prompt.format(
            file_name=os.path.basename(chunk["file_path"]),
            name=chunk["name"],
            block_type=chunk["block_type"],
            code=chunk["content"],
        )
        safe_config = GenerationConfig(temperature=0.0, max_output_tokens=4096)
        raw = self._generate_with_retries(prompt, config=safe_config)
        cleaned = raw.replace("```plantuml", "").replace("```puml", "").replace("```", "").strip()
        if not cleaned:
            raise UMLGenerationError(f"Manifest extraction produced empty content for {chunk['name']}", retryable=False)
        return cleaned

    def reduce_manifests_to_uml(self, file_name: str, manifests: List[str]) -> str:
        if not manifests:
            raise UMLGenerationError(f"No manifests found for file '{file_name}'", retryable=False)

        combined_manifests = "\n---\n".join(manifests)
        prompt = self.reduce_prompt.format(
            file_name=file_name,
            manifests=combined_manifests,
        )
        raw = self._generate_with_retries(prompt, config=self.generation_config)
        match = re.search(r"(@startuml[\s\S]*?@enduml)", raw)
        puml = match.group(1).strip() if match else raw
        sanitized = self._sanitize_puml(puml)
        if not sanitized:
            raise UMLGenerationError(f"Reduced UML content is empty for file '{file_name}'", retryable=False)
        return sanitized

    def _sanitize_puml(self, puml: str) -> str:
        puml = puml.replace("```plantuml", "").replace("```puml", "").replace("```", "")
        lines = [line.rstrip() for line in puml.splitlines() if line.strip()]
        if not lines:
            return ""
        if not lines[0].startswith("@startuml"):
            lines.insert(0, "@startuml")
        if not lines[-1].startswith("@enduml"):
            lines.append("@enduml")
        return "\n".join(lines)

    def generate_e2e_business_flow(self, call_graph_csv: str) -> str:
        prompt = f"""
You are an expert VB.NET Architect.
I am providing you with the complete Call Graph (caller to callee) of a VB.NET COM Server application in CSV format.

<CALL_GRAPH>
{call_graph_csv}
</CALL_GRAPH>

Your task is to analyze this call graph and generate a meaningful PlantUML Component Diagram that explains the true "End-to-End" workflow of the system.

CRITICAL RULES:
1. Define the core components (e.g. `component "Proxy" as P1`).
2. YOU MUST DRAW ARROWS (`-->`) BETWEEN COMPONENTS based on the CSV. A diagram with standalone boxes and no arrows is a failure.
3. Trace the logical workflow from the entry points (Proxies/Managers) down to the data access layer and DB.
4. Use `left to right direction`.
5. Provide a single `@startuml` ... `@enduml` block.
6. Do not use markdown wrappers.
"""
        raw = self._generate_with_retries(prompt, config=self.generation_config)
        match = re.search(r"(@startuml[\s\S]*?@enduml)", raw)
        puml = match.group(1).strip() if match else raw
        sanitized = self._sanitize_puml(puml)
        if not sanitized:
            raise UMLGenerationError("Failed to generate a valid E2E UML block", retryable=False)
        return sanitized
