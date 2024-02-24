using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HospitalLists
{
    public class SingletonClass
    {
        // Переменная instance является статической, чтобы была возможность обращаться к ней без создания объекта класса.
        private static SingletonClass instance;
        // Переменная autorizType хранит тип авторизации пользователя.
        private int autorizType;

        // конструктор класса является приватным, чтобы нельзя было создать объект класса извне
        private SingletonClass()
        {
            autorizType = -1; // устанавливаем значение по умолчанию
        }

        // метод для получения единственного объекта класса
        public static SingletonClass getInstance()
        {
            // Проверяем, создан ли объект класса
            if (instance == null)
            {
                // Если нет, то создаем его
                instance = new SingletonClass();
            }
            // Возвращаем единственный объект класса
            return instance;
        }

        // Метод doSomething() представляет собой пример метода, который будет выполняться на единственном объекте класса.
        public void doSomething()
        {
            // Реализация метода
        }

        // Метод для получения значения переменной autorizType
        public int getField1()
        {
            return autorizType;
        }

        // Метод для установки значения переменной autorizType
        public void setField1(int field1)
        {
            this.autorizType = field1;
        }

    }

    // Класс UserClass представляет собой класс, описывающий пользователя.
    public class UserClass
    {
        // Поле typeStaticInt является статическим и доступным для всех объектов данного класса.
        public static int typeStaticInt = 0;
        internal string User { get; set; }
        internal string Pass { get; set; }
        internal string FIO { get; set; }
        internal int Autoriz { get; set; }

        // Конструктор класса, принимающий тип авторизации
        public UserClass(int autoriz)
        {
            Autoriz = autoriz;
        }

        // Конструктор класса, принимающий ФИО пользователя
        public UserClass(string fio)
        {
            FIO = fio;
        }
    }
}
