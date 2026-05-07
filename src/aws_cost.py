import boto3
from datetime import datetime
from dateutil.relativedelta import relativedelta


def get_client(session):
    return session.client("ce")


# ═══════════════════════════════════════════
# Filter — Exclude Tax, Credits, Refunds
# Used for "Excluding Tax" tables (Billing view)
# ═══════════════════════════════════════════

BILLING_FILTER = {
    "Not": {
        "Dimensions": {
            "Key": "RECORD_TYPE",
            "Values": ["Tax", "Credit", "Refund"]
        }
    }
}


def get_monthly_cost(session, start, end, cost_filter=None):
    """
    Fetch total monthly cost.
    Pass cost_filter=BILLING_FILTER to exclude tax.
    Pass cost_filter=None for full Cost Explorer cost.
    """

    client = get_client(session)

    params = {
        "TimePeriod": {
            "Start": start,
            "End": end
        },
        "Granularity": "MONTHLY",
        "Metrics": ["UnblendedCost"]
    }

    if cost_filter:
        params["Filter"] = cost_filter

    response = client.get_cost_and_usage(**params)

    results = response.get("ResultsByTime", [])

    if not results:
        return 0.0

    amount = results[0] \
        ["Total"]["UnblendedCost"]["Amount"]

    return float(amount)


# Last 5 Months History
def get_last_5_months_cost(session, cost_filter=None):
    """
    Fetch cost for last 5 completed months.
    Pass cost_filter=BILLING_FILTER to exclude tax.
    Pass cost_filter=None for full Cost Explorer cost.
    """

    client = get_client(session)

    today = datetime.today()

    history = []

    for i in range(5, 0, -1):

        start = (
            today.replace(day=1)
            - relativedelta(months=i)
        )

        end = (
            today.replace(day=1)
            - relativedelta(months=i - 1)
        )

        params = {
            "TimePeriod": {
                "Start": start.strftime("%Y-%m-%d"),
                "End": end.strftime("%Y-%m-%d")
            },
            "Granularity": "MONTHLY",
            "Metrics": ["UnblendedCost"]
        }

        if cost_filter:
            params["Filter"] = cost_filter

        response = client.get_cost_and_usage(**params)

        results = response.get("ResultsByTime", [])

        if results:
            amount = float(
                results[0]
                ["Total"]["UnblendedCost"]["Amount"]
            )
        else:
            amount = 0.0

        history.append({
            "Month": start.strftime("%b %Y"),
            "Cost": amount
        })

    return history