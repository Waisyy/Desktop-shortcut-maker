import os
import win32com.client

def create_shortcut(exe_file, shortcut_path):
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.Targetpath = exe_file
    shortcut.Save()

def find_exe_files(directory):
    exe_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith(".exe"):
                exe_files.append(os.path.join(root, file))
    return exe_files

def main():
    directory = input("Введите путь к папке: ")
    if os.path.exists(directory):
        exe_files = find_exe_files(directory)
        if exe_files:
            print("Найденные .exe файлы:")
            for file in exe_files:
                print(file)
            answer = input("Хотите создать ярлыки на рабочий стол? (да/нет): ")
            if answer.lower() == "да":
                desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
                for file in exe_files:
                    answer = input(f"Является ли {file} игрой? (да/нет): ")
                    if answer.lower() == "да":
                        filename = os.path.basename(file)
                        shortcut_path = os.path.join(desktop, filename + ".lnk")
                        if not os.path.exists(shortcut_path):
                            create_shortcut(file, shortcut_path)
                            print(f"Создан ярлык для {filename} на рабочем столе")
                        else:
                            print(f"Ярлык для {filename} уже существует на рабочем столе")
                    else:
                        print(f"{file} не является игрой")
        else:
            print("Не найдено .exe файлов в указанной папке")
    else:
        print("Указанная папка не существует")

if __name__ == "__main__":
    main()