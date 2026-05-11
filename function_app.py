import logging
import azure.functions as func

from Timesheethour_main import generate_weekly_timesheet

app = func.FunctionApp()


@app.schedule(
    schedule="0 0 7 * * TUE",
    arg_name="myTimer",
    run_on_startup=False,
    use_monitor=True
)
def weekly_timesheet_updater(myTimer: func.TimerRequest) -> None:

    logging.info("Weekly timesheet updater started.")

    generate_weekly_timesheet()

    logging.info("Weekly timesheet updater completed.")