from validator.BaseValidator import BaseValidator
import pycountry


class CountryValidator(BaseValidator):

    message = "This value is not correct country name"
    countries = pycountry.countries

    def validate(self, value):
        # Allow empty values
        if value is None:
            return True

        value = super().validate(value)

        try:
            CountryValidator.countries.get(name=value)
            return True
        except KeyError:
            return False

    def __init__(self, params):
        super().__init__(params)