from abc import ABC, abstractmethod


class BaseValidator(ABC):
    trim = False

    @property
    @abstractmethod
    def message(self):
        pass

    @abstractmethod
    def validate(self, value):
        if self.trim and isinstance(value, str):
            return value.strip()

        return value

    def getMessage(self):
        return self.message

    def __init__(self, params):
        if params:
            if "message" in params:
                self.message = params["message"]

            if "trim" in params:
                self.trim = params["trim"]