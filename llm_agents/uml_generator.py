import re
from typing import List, Dict
from vertexai.generative_models import GenerativeModel, GenerationConfig, SafetySetting

class UMLGenerator:
    def __init__(self, model_name: str, temperature: float = 0.0, max_tokens: int = 4096):
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
You are an expert VB.NET developer and Architect.
Your task is to generate a PlantUML Activity Diagram that explicitly details the computational flow of the following VB.NET code block.
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

Instructions:
1. Parse the TARGET_CODE logic into a sequence of steps, decisions (if/else), and loops.
2. If TARGET_CODE calls another method, cleanly define that as an activity node.
3. Write ONLY the PlantUML string starting with `@startuml` and ending with `@enduml`. DO NOT use markdown code blocks like ```plantuml, just output the raw text.

Example format:
@startuml
start
:Initialize Variables;
if (Condition is true?) then (yes)
  :Process Data;
else (no)
  :Log Error;
endif
stop
@enduml
"""

    def process_chunk(self, chunk: Dict, context_chunks: List[Dict]) -> str:
        context_str = "\n".join([f"--- Name: {c['name']} ---\n{c['content']}" for c in context_chunks])
        prompt = self.prompt_template.format(
            context=context_str,
            name=chunk['name'],
            file_path=chunk['file_path'],
            block_type=chunk['block_type'],
            code=chunk['content']
        )

        import time
        max_retries = 3
        for attempt in range(max_retries):
            try:
                response = self.model.generate_content(
                    prompt,
                    generation_config=self.generation_config,
                    safety_settings=self.safety_settings
                ).text
                
                match = re.search(r"(@startuml[\s\S]*?@enduml)", response)
                if match:
                    return match.group(1).strip()
                return response.strip()
                
            except Exception as e:
                err_str = str(e)
                if '429' in err_str and attempt < max_retries - 1:
                    wait = 5 * (2 ** attempt)
                    time.sleep(wait)
                    continue
                elif '429' not in err_str:
                    return ""
        return ""

    def generate_macro_uml(self, all_names: List[str]) -> str:
        macro_prompt = f"""
You are an expert VB.NET Architect.
I have parsed a large codebase and extracted the following component names (Classes, Modules, Methods):
{all_names}

Your task is to generate a SINGLE top-level PlantUML Sequence/Flow Diagram that represents the complete end-to-end design flow.
Use the components provided to trace a logical system execution sequence instead of just a static class map, showing how they call one another.
Output ONLY the raw PlantUML text starting with @startuml and ending with @enduml.
CRITICAL: Do NOT write any PlantUML comments (lines starting with `'`). Do NOT add any extra explanation text.
Keep the diagram clean, simple, and strictly valid PlantUML syntax.
"""
        try:
            response = self.model.generate_content(macro_prompt, generation_config=self.generation_config).text
            match = re.search(r"(@startuml[\s\S]*?@enduml)", response)
            
            if match:
                res = match.group(1).strip()
            else:
                res = response.replace("```plantuml", "").replace("```puml", "").replace("```", "").strip()
                if not res.startswith("@startuml"):
                    res = "@startuml\n" + res
                if not res.endswith("@enduml"):
                    res = res + "\n@enduml"
            
            res = res.replace("module ", "class ")
            return res
        except:
            return ""
