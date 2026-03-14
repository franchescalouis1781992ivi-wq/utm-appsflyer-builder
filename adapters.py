from dataclasses import dataclass


@dataclass(frozen=True)
class LinkAdapter:
    name: str
    provider: str
    phase: str = "phase_1"

    def describe(self) -> str:
        if self.provider == "local_passthrough":
            return "当前仅做本地参数拼接与 AppsFlyer 表格生成，不调用外部 API。"
        return "预留给后续外部转链能力。"

    def transform(self, web_url: str, app_deeplink: str) -> dict[str, str]:
        return {
            "deep_link_value": app_deeplink,
            "af_dp": app_deeplink,
            "af_web_dp": web_url,
            "af_android_url": web_url,
            "af_ios_url": web_url,
        }


ADAPTER_REGISTRY = {
    "local_passthrough": LinkAdapter(
        name="local_passthrough",
        provider="local_passthrough",
    ),
    "appsflyer_onelink": LinkAdapter(
        name="appsflyer_onelink",
        provider="appsflyer",
        phase="future",
    ),
    "airbridge": LinkAdapter(
        name="airbridge",
        provider="airbridge",
        phase="future",
    ),
    "branch": LinkAdapter(
        name="branch",
        provider="branch",
        phase="future",
    ),
}
