import os
import logging
from typing import Tuple, List, Dict
import openai

logger = logging.getLogger(__name__)

class RAGService:
    def __init__(self):
        # In a real impl, initialize Qdrant client here
        self.api_key = os.environ.get("OPENAI_API_KEY")
        if self.api_key:
            self.client = openai.AsyncOpenAI(api_key=self.api_key)
        else:
            logger.warning("OPENAI_API_KEY not set. LLM generation will be mocked.")
            self.client = None

    async def answer(self, query: str) -> Tuple[str, List[Dict]]:
        """
        Retrieve context and generate answer.
        
        Returns:
            (answer_text, list_of_context_sources)
        """
        
        # 1. Retrieve (Mock for PoC - Real implementation needs Qdrant connection)
        # TODO: Connect to Qdrant and search 'pptx_slides' collection
        context_docs = self._mock_retrieve(query)
        
        context_text = "\n\n".join([f"Source: {d['file_name']}\nContent: {d['text']}" for d in context_docs])
        
        # 2. Generate
        if self.client:
            try:
                prompt = f"""
                You are a helpful assistant answering based ONLY on the provided context.
                
                Context:
                {context_text}
                
                Question: {query}
                
                Answer:
                """
                
                response = await self.client.chat.completions.create(
                    model="gpt-3.5-turbo",
                    messages=[
                        {"role": "system", "content": "You are a helpful assistant."},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.0
                )
                answer = response.choices[0].message.content
            except Exception as e:
                logger.error(f"OpenAI API error: {e}")
                answer = "Error communicating with LLM service."
        else:
            answer = f"[Mock LLM] I received your query: '{query}'.\nContext found: {len(context_docs)} docs."

        return answer, context_docs

    def _mock_retrieve(self, query: str) -> List[Dict]:
        """Mock retrieval for basic bot testing without full Qdrant setup."""
        return [
            {
                "file_name": "design_guide_v1.pptx",
                "slide_no": 5,
                "text": "The hinge mechanism requires a tolerance of +/- 0.1mm."
            },
            {
                "file_name": "thermal_analysis.pptx",
                "slide_no": 12,
                "text": "Max operating temperature is 85 degrees Celsius."
            }
        ]
