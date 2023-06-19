from botcity.plugins.telegram import BotTelegramPlugin
from botcity.core import DesktopBot
import telebot
from telebot.apihelper import ApiTelegramException

class Bot(DesktopBot):
    def action(self, execution=None):



        telegram = BotTelegramPlugin(token='5825853452:AAH77EiouxZreHvnd7p_6DsAlRDCCdB2cbk')
        try:
            response = telegram.send_message(
                text="Mantendo Conectado",
                group="EnvioMensagens",
                username=["@Rafael_25Bot"]
            )

        except ApiTelegramException as e:
            if e.error_code == 400 and "chat not found" in e.description:
                # Trate o erro de chat não encontrado aqui
                print("Erro: Chat não encontrado!")
                pass
            #else:
            #    # Trate outros erros que você possa querer lidar de forma diferente
            #    print('PASSOU')
            #    print(f"Erro: {e}")

        print('segue a vida')




    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()
