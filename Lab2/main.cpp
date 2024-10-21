#include "Models/Excel.h"
#include <cassert>


int main(){

    /*Выполнение 1 лабораторной работы и её тестирование*/

    ValueCell A1;
    assert(A1.getNumericValue() == 0);

    ValueCell A2{1425}; // Теперь хранит только числовое значение
    assert(A2.getNumericValue() == 1425);

    ValueCell A3{"Кукла"};
    assert(A3.getTextValue() == "Кукла");

    ValueCell A4{123}; // Хранит числовое значение
    assert(A4.getNumericValue() == 123);

    // Проверка на исключения
    try {
        A3.getNumericValue(); // Это должно вызвать исключение
    } catch (const runtime_error& e) {
        cout << e.what() << endl; // Ожидаем сообщение об ошибке
    }

    // Проверка на установку значений
    A2.setNumericValue(321);
    assert(A2.getNumericValue() == 321);

    A2.setTextValue("Ромашки"); // Это изменит тип ячейки на текстовое значение
    assert(A2.getTextValue() == "Ромашки");

    // Проверка конструктора копирования
    ValueCell A5 = A2; // Используем конструктор копирования
    assert(A5.getTextValue() == "Ромашки");

    // Попробуем получить числовое значение из A5
    try {
        A5.getNumericValue(); // Это должно вызвать исключение
    } catch (const runtime_error& e) {
        cout << e.what() << endl; // Ожидаем сообщение об ошибке
    }

    cout << "Все тесты прошли проверку!" << endl;



    /*Выполнение 2 лабораторной работы и её тестирование*/

     int numRows, numCols;
    cout << "Введите количество строк и столбцов: ";
    cin >> numRows >> numCols;

    vector<vector<ValueCell>> values(numRows, vector<ValueCell>(numCols));

    // Заполнение матрицы значениями
    for (int i = 0; i < numRows; ++i) {
        for (int j = 0; j < numCols; ++j) {
            string input;
            cout << "Введите числовое значение для ячейки (" << i << ", " << j << "): ";
            cin >> input;
            try {
                double numericValue = stod(input);
                values[i][j] = ValueCell(numericValue); // Сохраняем как числовое значение
            } catch (invalid_argument&) {
                values[i][j] = ValueCell(input); // Сохраняем как текстовое значение
            }
        }
    }


    // Вывод матрицы
    FormulaCell dummy(0, 0, numRows - 1, numCols - 1, values); // Временная ячейка для вывода
    dummy.printMatrix();

    // Запрос диапазона у пользователя
    int startRow, startCol, endRow, endCol;
    cout << "Введите диапазон для операций" << endl;
    cout << "Введите начальную строку: ";
    cin >> startRow;
    cout << "Введите начальный столбец: ";
    cin >> startCol;
    cout << "Введите конечную строку: ";
    cin >> endRow;
    cout << "Введите конечный столбец: ";
    cin >> endCol;

    // Создание формульной ячейки с заданным диапазоном
    FormulaCell formula(startRow, startCol, endRow, endCol, values);

    cout << "Сумма: " << formula.sum() << endl;
    cout << "Произведение: " << formula.product() << endl;
    cout << "Среднее: " << formula.average() << endl;

    formula.printRange();

    return 0;
}