from dotenv import load_dotenv
load_dotenv('config/.env')
import logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(message)s')
from jobsignal.agents.notify_bot import NotifyBot
NotifyBot().run_webhook_polling()
