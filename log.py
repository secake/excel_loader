import datetime


class Log:
    class Level:
        DEBUG = 0
        INFO = 1
        WARN = 2
        ERROR = 3

        def str(level):
            return {
                Log.Level.DEBUG: 'DEBUG',
                Log.Level.INFO: 'INFO',
                Log.Level.WARN: 'WARN',
                Log.Level.ERROR: 'ERROR',
            }.get(level, 'ERROR')

    import io
    def __init__(self, level=Level.DEBUG, output=io.StringIO()):
        self.level = level
        self.output = output

    def read(self):
        logs = ''
        if self.output:
            self.output.seek(0)
            logs = self.output.read()
            self.output.flush()
        return logs

    def log(self, level, *args):
        form_msg = ''
        if self.level <= level:
            form_msg = '{time} 【{level}】: {msg}'.format(
                time=datetime.datetime.now().strftime(
                    '%Y-%m-%d %H:%M:%S'),
                level=Log.Level.str(level),
                msg=' '.join([str(arg) for arg in args])
            )
            print(form_msg)
            if self.output:
                self.output.write(form_msg+'\n')
        return form_msg

    def err(self, *args):
        return self.log(Log.Level.ERROR, *args)

    def warn(self, *args):
        return self.log(Log.Level.WARN, *args)

    def info(self, *args):
        return self.log(Log.Level.INFO, *args)

    def debug(self, *args):
        return self.log(Log.Level.DEBUG, *args)

