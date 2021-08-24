import json
import re
from time import sleep
from os import system


def clear():
    system('cls')


class EditDictMenu:
    json_path = '../options/contact_info.json'
    value_list = ['Имя', 'Фамилия', 'Телефон', 'Почта', 'Метро', 'Улица', 'Дом', 'Стоение', 'Корпус', 'Подьезд', 'Этаж',
                  'Квартира', 'Домофон']
    key_list = []

    def __init__(self):
        try:
            with open(self.json_path, 'r', encoding='utf-8') as json_file:
                self.key_list = json.load(json_file)
        except Exception as e:
            print(e.args)

    def save(self):
        with open(self.json_path, 'w', encoding='utf-8') as json_file:
            json.dump(self.key_list, json_file)

    def show_dict(self):
        for i in range(0, len(self.key_list)):
            dct = self.key_list[i]
            print(f'{i+1}):', end='')
            for key in self.key_list[i].keys():
                print(f' {key}: {self.key_list[i][key]}', end='')
            print(' ')

    def add_dict(self):
        print('Добавить новую контактную информацию:')
        result_dict = {}
        for value in self.value_list:
            result_dict.update({value: input(f'{value}: ')})
        self.key_list.append(result_dict)

    def delete_dict(self, num):
        if len(self.key_list) >= num >= 1:
            self.key_list.pop(num-1)
        else:
            print('Нет контактной информации с таким номером!')
            sleep(2)

    def run(self):
        loop = True
        while loop:
            self.show_dict()
            command = input('Commands:\n '
                            'add: добавить новую контактную информацию\n '
                            'delete <num>: удалить контактную информацию\n '
                            's: сохранить\n '
                            'q: выйти\n')
            if command == 'add':
                self.add_dict()
            elif re.match('delete', command):
                try:
                    self.delete_dict(int(re.split(' ', command)[1]))
                except Exception as e:
                    print(e)
            elif command == 's':
                self.save()
            elif command == 'q':
                self.save()
                loop = False
            clear()


edit = EditDictMenu()
edit.run()
