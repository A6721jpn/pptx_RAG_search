from botbuilder.core import ActivityHandler, TurnContext
from botbuilder.schema import ChannelAccount
import logging
from src.rag.generator import RAGService

logger = logging.getLogger(__name__)

class RAGBot(ActivityHandler):
    def __init__(self):
        self.rag_service = RAGService()

    async def on_members_added_activity(
        self, members_added: ChannelAccount, turn_context: TurnContext
    ):
        for member_added in members_added:
            if member_added.id != turn_context.activity.recipient.id:
                await turn_context.send_activity("Hello! I am the PPTX RAG Bot. Ask me anything about the documents (English only).")

    async def on_message_activity(self, turn_context: TurnContext):
        query = turn_context.activity.text.strip()
        
        # Call RAG Service
        try:
            # Indicate typing
            # await turn_context.send_activity(Activity(type=ActivityTypes.typing))
             
            answer, context_docs = await self.rag_service.answer(query)
            
            await turn_context.send_activity(answer)
            
            # Optionally show sources
            if context_docs:
                sources_text = "\n\n**Sources:**\n" + "\n".join([f"- {doc['file_name']} (Slide {doc['slide_no']})" for doc in context_docs])
                await turn_context.send_activity(sources_text)
                
        except Exception as e:
            logger.error(f"Error generating answer: {e}")
            await turn_context.send_activity("Sorry, I encountered an error while processing your request.")
