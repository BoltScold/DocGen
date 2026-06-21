import azure.functions as func
import datetime
import logging
from sharepoint_job import main

app = func.FunctionApp()

@app.timer_trigger(
    schedule="0 0 20 * * *",  # 20:00 daily
    arg_name="myTimer",
    run_on_startup=False,
    use_monitor=True  # ✅ change this!
)
def MyScheduledScript(myTimer: func.TimerRequest) -> None:
    
    if myTimer.past_due:
        logging.warning('Timer is past due!')

    logging.info('Timer trigger started')

    try:
        main()  # ✅ your script runs here
        logging.info('Timer trigger completed successfully')

    except Exception as e:
        logging.error(f"Error during execution: {e}")
        raise