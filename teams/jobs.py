from datetime import date, datetime

from django.conf import settings
import logging
logger = logging.getLogger('django')

def schedule_api():
    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    logger.info("apitesting"+ dt_string)
    print("api" + str(dt_string))

