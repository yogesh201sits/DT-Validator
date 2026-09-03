from validator.BaseValidator import BaseValidator


class NumericRangeValidator(BaseValidator):

    message = "This value is outside the allowed numeric range"

    def validate(self, value):
        if value is None:
            return True

        value = super(NumericRangeValidator, self).validate(value)

        try:
            number = float(value)
        except (TypeError, ValueError):
            return False

        if self.minimum is not None and number < self.minimum:
            return False

        if self.maximum is not None and number > self.maximum:
            return False

        return True

    def __init__(self, params):
        super(NumericRangeValidator, self).__init__(params)
        self.minimum = params.get("min")
        self.maximum = params.get("max")

        if self.minimum is None and self.maximum is None:
            raise ValueError("NumericRange requires min or max")

        if self.minimum is not None:
            self.minimum = float(self.minimum)
        if self.maximum is not None:
            self.maximum = float(self.maximum)
        if (
            self.minimum is not None
            and self.maximum is not None
            and self.minimum > self.maximum
        ):
            raise ValueError("NumericRange min cannot exceed max")
