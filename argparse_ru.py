import gettext

def localization(Text):
    Text = Text.replace("usage", "Применение")
    Text = Text.replace("show this help message and exit",
                        "выводит это сообщение и завершает работу")
    Text = Text.replace("error:", "Ошибка:")
    Text = Text.replace("the following arguments are required:",
                        "Следующие аргументы обязательные:")
    Text = Text.replace("options",
                        "Параметры")
    Text = Text.replace("show program's version number and exit",
                        "Показывает версию и завершает работу")
    Text = Text.replace("unrecognized arguments",
                        "нераспознанные параметры")
    Text = Text.replace("examples:",
                        "Примеры:")
    return Text
gettext.gettext = localization