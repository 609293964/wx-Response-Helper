class WeChatLocale:
    MAPPING = {
        "weixin": {"en-US": "Weixin", "zh-CN": "微信", "zh-TW": "微信"},
        "message": {"en-US": "消息", "zh-CN": "消息", "zh-TW": "消息"},
    }

    def __init__(self, locale="zh-CN"):
        for key, value in WeChatLocale.MAPPING.items():
            setattr(self, key, value[locale])

    @staticmethod
    def getSupportedLocales():
        return list(WeChatLocale.MAPPING["weixin"].keys())
