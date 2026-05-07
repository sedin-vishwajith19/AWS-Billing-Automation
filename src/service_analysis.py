import boto3
from datetime import datetime
from dateutil.relativedelta import relativedelta


def get_service_costs(session, start, end, cost_filter=None):
    """
    Fetch per-service cost breakdown.
    Pass cost_filter to exclude tax/credits/refunds.
    Handles pagination for complete results.
    """

    client = session.client("ce")

    services = {}
    next_token = None

    while True:

        params = {
            "TimePeriod": {
                "Start": start,
                "End": end
            },
            "Granularity": "MONTHLY",
            "Metrics": ["UnblendedCost"],
            "GroupBy": [
                {
                    "Type": "DIMENSION",
                    "Key": "SERVICE"
                }
            ]
        }

        if cost_filter:
            params["Filter"] = cost_filter

        if next_token:
            params["NextPageToken"] = next_token

        response = client.get_cost_and_usage(**params)

        for period in response.get("ResultsByTime", []):

            for group in period.get("Groups", []):

                name = group["Keys"][0]

                cost = float(
                    group["Metrics"]
                    ["UnblendedCost"]["Amount"]
                )

                # Accumulate cost if service appears
                # across multiple pages
                services[name] = (
                    services.get(name, 0) + cost
                )

        next_token = response.get("NextPageToken")

        if not next_token:
            break

    return services


def find_primary_reason(prev, curr):

    all_services = set(
        list(prev.keys()) + list(curr.keys())
    )

    changes = {}

    for service in all_services:

        prev_cost = prev.get(service, 0)
        curr_cost = curr.get(service, 0)

        change = curr_cost - prev_cost

        if abs(change) > 0.01:
            changes[service] = round(change, 2)

    # Sort by largest absolute change
    sorted_services = sorted(
        changes.items(),
        key=lambda x: abs(x[1]),
        reverse=True
    )

    return sorted_services[:3]