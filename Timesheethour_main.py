from ExtraUpdateSupplier import run_update_supplier
from Timesheethours1 import run_weekly_replace_roster_hours
from Timesheethours2 import generate_weekly_timesheet
from TimesheethourEmail import send_email_with_attachment


def main():

    # ==================================================
    # 1. Update Supplier First
    # ==================================================

    run_update_supplier()

    # ==================================================
    # 2. Update SharePoint roster hours
    # ==================================================

    run_weekly_replace_roster_hours()

    # ==================================================
    # 3. Generate Excel report in memory
    # ==================================================

    (
        excel_bytes,
        output_file,
        week_start,
        week_end
    ) = generate_weekly_timesheet()

    # ==================================================
    # 4. Send email
    # ==================================================

    subject = (
        f"Weekly Timesheet Report "
        f"{week_start.strftime('%Y/%m/%d')} - "
        f"{week_end.strftime('%Y/%m/%d')}"
    )

    body = (
        "Hi Planning Team,\n\n"
        "Please find attached the weekly timesheet report.\n\n"
        "Regards,\n"
        "Automated Timesheet System"
    )

    send_email_with_attachment(
        subject,
        body,
        excel_bytes,
        output_file
    )


if __name__ == "__main__":
    main()