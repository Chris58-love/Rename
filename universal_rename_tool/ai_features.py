from .ai_client import AIConfig, DSV4FlashClient


def build_ai_config(enabled=False, api_key="", base_url="", model="", timeout=30, temperature=0.2):
    return AIConfig(
        enabled=bool(enabled),
        api_key=api_key or "",
        base_url=base_url or "",
        model=model or "dsv4flash",
        timeout=int(timeout or 30),
        temperature=float(temperature if temperature is not None else 0.2),
    )


def test_ai_connection(config: AIConfig):
    return DSV4FlashClient(config).test_connection()
