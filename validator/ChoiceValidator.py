from validator.BaseValidator import BaseValidator


class ChoiceValidator(BaseValidator):

    message = "This value is not valid choice"
    choices = []
    caseSensitive = True

    def validate(self, value):
        # Allow empty values
        if value is None:
            return True

        value = super().validate(value)

        if not self.caseSensitive:
            value = value.lower()

        if value in self.choices:
            return True

        return False

    def __init__(self, params):
        super().__init__(params)

        if "choices" not in params:
            raise ValueError("Valid choice are not set")

        self.choices = params.get("choices")

        if "caseSensitive" in params:
            self.caseSensitive = params.get("caseSensitive")