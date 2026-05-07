import os
import yaml


def load_accounts():

    env_accounts = os.environ.get("AWS_ACCOUNTS")

    if env_accounts:

        accounts = []

        for entry in env_accounts.split(","):

            entry = entry.strip()

            if not entry:
                continue

            parts = entry.split(":")

            if len(parts) == 3:
                accounts.append({
                    "name": parts[0].strip(),
                    "id": parts[1].strip(),
                    "profile": parts[2].strip()
                })
            elif len(parts) == 4:
                accounts.append({
                    "name": parts[0].strip(),
                    "id": parts[1].strip(),
                    "access_key": parts[2].strip(),
                    "secret_key": parts[3].strip()
                })
            else:
                print(
                    f"⚠ Skipping invalid account entry: "
                    f"'{entry}' — expected name:id:profile OR name:id:access_key:secret_key"
                )
                continue

        if not accounts:
            raise ValueError(
                "AWS_ACCOUNTS env variable is set but "
                "contains no valid entries."
            )

        print(
            f"✔ Loaded {len(accounts)} account(s) "
            f"from AWS_ACCOUNTS env variable"
        )

        return accounts

    # Fallback to YAML file
    with open("config/accounts.yaml", "r") as file:
        config = yaml.safe_load(file)

    print(
        f"✔ Loaded {len(config['accounts'])} account(s) "
        f"from config/accounts.yaml"
    )

    return config["accounts"]