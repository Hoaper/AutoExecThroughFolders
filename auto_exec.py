from colorama import Fore, Style
from subprocess import Popen, PIPE
from openpyxl import Workbook
from time import strftime, sleep
import os

# Clear terminal
os.system('cls')

# Vars init
path = ""
exec_name = ""
delay = 5

def get_info():

    # Get directory
    path = input(f"Введите путь к папкам {Fore.YELLOW}(По умолчанию: {os.getcwd()}){Style.RESET_ALL} \n")
    path = path if len(path) > 0 else os.getcwd()

    # Get name from stdin
    exec_name = input(f"\nИмя исполняемого файла{Style.RESET_ALL}\n")

    # Delay settings
    try:
        delay = int(input(f"\nЗадержка между запуском приложения в секундах {Fore.YELLOW}(По умолчанию: 5 секунд){Style.RESET_ALL}\n"))
    except Exception:
        delay = 5

    print(
            f"""
            Пожалуйста, проверьте правильность введенных данных:
            Путь к папкам: {Fore.YELLOW}{path}
            {Style.RESET_ALL}Исполняемый файл: {Fore.YELLOW}{exec_name}
            {Style.RESET_ALL}Задержка между запуском: {Fore.YELLOW}{delay} секунд
            """
        )
    def ask_sure():
        answer = input(f"{Fore.MAGENTA}Все верно?{Style.RESET_ALL} (y/n) ")
        if answer in ["y", "Y"]:
            return path, exec_name, delay
        elif answer in ["n", "N"]:
            return get_info()
        else:
            return ask_sure()

    return ask_sure()

try:
    path, exec_name, delay = get_info()
except KeyboardInterrupt:
    print("Завершаю работу...")
    input("Нажмите ENTER чтобы выйти...")
# Counting processes
num = 1
fail = 0

# Excel output init
wb = Workbook()
ws = wb.active
ws.append(["НОМЕР", "ПУТЬ К ФАЙЛУ", "РЕЗУЛЬТАТ"])

def save_excel():
    print(f"Сохраняем Excel таблицу... ")
    excel_path = strftime(r"%H%M%S_%m%d%Y")+'.xlsx'

    wb.save(excel_path)
    print(f"Сохранено как: {excel_path}")


# Starting programs part
try:
    for e in [f.path for f in os.scandir(path) if f.is_dir()]:
        e_p = os.path.join(path, e, exec_name)
        try:
            print(f"{num}. {Fore.YELLOW}{e_p}{Style.RESET_ALL}... ", end="")
            err_code = Popen([e_p], stdout=PIPE, stderr=PIPE)
        except FileNotFoundError:
            print(f"{Fore.RED}не найден{Style.RESET_ALL}")
            ws.append([num-1, e_p, "не найден"])
            fail += 1
            num += 1
            continue

        except Exception:
            print(f"{Fore.RED}провал", end=Style.RESET_ALL+"\n")
            print("Ошибка выполнения данного файла")
            ws.append([num-1, e_p, "провал запуска"])
            fail += 1
            num += 1
            continue

        print(f"{Fore.MAGENTA}успех", end=Style.RESET_ALL+"\n")
        num += 1
        ws.append([num-1, e_p, "УСПЕХ"])
        sleep(delay)

except KeyboardInterrupt:
    print(f"\nВы прервали выполнение программы!")
    def ask_save():
        ask = input(f"Сохранить Excel файл? (y/n) ")
        if ask in ['y', 'Y']:

            save_excel()
        elif ask in ['n', 'N']:
            return
        else:
            ask_save()
    ask_save()
    input("Нажмите ENTER чтобы выйти...")
    exit()

except Exception as e:
    print(e)
    print("Данной директории не существует или в ней нет папок, убедитесь, что все корректно введено")
    input("Нажмите ENTER чтобы выйти...")
    exit()

# Ending part
print(f"\nВсего файлов: {num-1}\nУспех {num-fail-1} | Провал {fail}")

save_excel()
input("Нажмите ENTER чтобы выйти...")
