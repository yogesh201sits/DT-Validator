from urllib.parse import urlparse

from validator.BaseValidator import BaseValidator


class URLValidator(BaseValidator):

    message = "This value is not a valid URL"

    def validate(self, value):
        if value is None:
            return True

        value = super(URLValidator, self).validate(value)

        if not isinstance(value, str):
            return False

        parsed = urlparse(value)
        return parsed.scheme in ("http", "https") and bool(parsed.netloc)

    def __init__(self, params):
        super(URLValidator, self).__init__(params)
