from botbuilder.core import Middleware, TurnContext
from botbuilder.schema import ActivityTypes
from langdetect import detect, LangDetectException
import logging

logger = logging.getLogger(__name__)

class LanguageValidationMiddleware(Middleware):
    """
    Middleware to ensure input is in English.
    """
    
    async def on_turn(self, context: TurnContext, logic):
        if context.activity.type == ActivityTypes.message and context.activity.text:
            user_text = context.activity.text
            
            try:
                # Simple check: Try to detect language
                # Only check if text is long enough to be confident? 
                # For very short texts (hi, hello), verification might be tricky.
                # But strict requirement is "English Only".
                
                lang = detect(user_text)
                
                if lang != 'en':
                    logger.info(f"Detected non-English language: {lang} for text '{user_text}'")
                    await context.send_activity("Error: Please use English only.")
                    return # Stop processing
                    
            except LangDetectException:
                # If cannot detect, might be code or mixed. Allow pass or strict fail?
                # Usually safely fail or warn. For now let's pass if unsure, or fail safer?
                # 'English only' -> fail safer? Let's allow if detection fails (often short text).
                pass
                
        # Proceed to next middleware/bot logic
        await logic()
