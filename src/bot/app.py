import sys
import traceback
import logging
from datetime import datetime

from aiohttp import web
from botbuilder.core import (
    BotFrameworkAdapter,
    BotFrameworkAdapterSettings,
    TurnContext,
    MiddlewareSet,
)
from botbuilder.schema import Activity, ActivityTypes

from src.bot.config import DefaultConfig
from src.bot.middleware.language import LanguageValidationMiddleware
from src.bot.bots.rag_bot import RAGBot

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

CONFIG = DefaultConfig()

# Create Adapter
SETTINGS = BotFrameworkAdapterSettings(CONFIG.APP_ID, CONFIG.APP_PASSWORD)
ADAPTER = BotFrameworkAdapter(SETTINGS)

# Add Middleware
ADAPTER.use(LanguageValidationMiddleware())

# Create Bot
BOT = RAGBot()

async def on_error(context: TurnContext, error: Exception):
    logger.error(f"\n [on_turn_error] unhandled error: {error}")
    traceback.print_exc()

    await context.send_activity("The bot encountered an error or bug.")
    await context.send_activity("To continue to run this bot, please fix the bot source code.")

ADAPTER.on_turn_error = on_error

async def messages(req: web.Request) -> web.Response:
    if "application/json" in req.headers["Content-Type"]:
        body = await req.json()
    else:
        return web.Response(status=415)

    activity = Activity().deserialize(body)
    auth_header = req.headers.get("Authorization", "")

    try:
        response = await ADAPTER.process_activity(activity, auth_header, BOT.on_turn)
        if response:
            return web.json_response(data=response.body, status=response.status)
        return web.Response(status=201)
    except Exception as exception:
        raise exception

APP = web.Application()
APP.router.add_post("/api/messages", messages)

if __name__ == "__main__":
    try:
        web.run_app(APP, host="localhost", port=CONFIG.PORT)
    except Exception as error:
        raise error
