from datetime import datetime

class LogMixim:

    @staticmethod
    def write(msg):
        with open('log.log', 'a+') as f:
            tempo = datetime.now().strftime('%d/%m/%Y %H:%M')
            f.write(f'{msg} - {tempo}\n')

    def log_info(self,msg):
        self.write(f'INFO: {msg}')
    
    def log_erro(self, msg):
        self.write(f'ERRO: {msg}')