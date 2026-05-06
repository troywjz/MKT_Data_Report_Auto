import os
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent


def load_env(path=".env"):
    env_path = BASE_DIR / path
    if not env_path.exists():
        return

    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue

        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


load_env()


def get_env(name, default=""):
    return os.getenv(name, default)


def get_int_env(name, default):
    value = os.getenv(name)
    if value in (None, ""):
        return default
    try:
        return int(value)
    except ValueError:
        return default


def get_platform_credentials(prefix):
    username = get_env(f"{prefix}_USERNAME")
    password = get_env(f"{prefix}_PASSWORD")
    if username or password:
        return username, password
    return None, None


def get_chromium_user_data_path():
    return get_env("CHROMIUM_USER_DATA_PATH", "./User_Data_Chrome")


def get_chromium_local_port():
    return get_int_env("CHROMIUM_LOCAL_PORT", 9222)
