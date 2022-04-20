class JsonError(Exception):
    def __init__(self, message):
        super(JsonError, self).__init__(message)
