from datetime import date, datetime

from validator.BaseValidator import BaseValidator


class DateRangeValidator(BaseValidator):

    message = "This value is outside the allowed date range"
    format = "%Y-%m-%d"

    def validate(self, value):
        if value is None:
            return True

        value = super(DateRangeValidator, self).validate(value)
        parsed_value = self._parse(value)

        if parsed_value is None:
            return False
        if self.minimum is not None and parsed_value < self.minimum:
            return False
        if self.maximum is not None and parsed_value > self.maximum:
            return False

        return True

    def _parse(self, value):
        if isinstance(value, datetime):
            return value.date()
        if isinstance(value, date):
            return value
        if isinstance(value, str):
            try:
                return datetime.strptime(value, self.format).date()
            except ValueError:
                return None
        return None

    def __init__(self, params):
        super(DateRangeValidator, self).__init__(params)
        self.minimum = self._parse(params.get("min")) if params.get("min") else None
        self.maximum = self._parse(params.get("max")) if params.get("max") else None

        if self.minimum is None and self.maximum is None:
            raise ValueError("DateRange requires min or max")
        if (
            self.minimum is not None
            and self.maximum is not None
            and self.minimum > self.maximum
        ):
            raise ValueError("DateRange min cannot exceed max")
