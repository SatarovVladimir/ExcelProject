#include "Excel.h"

ValueCell::ValueCell(): numericValue(0), textValue(""), isnumericValue(true) {}
ValueCell::ValueCell(double Numeric): numericValue(Numeric), textValue(""), isnumericValue(true){}
ValueCell::ValueCell(string Text): numericValue(0), textValue(Text), isnumericValue(false){}
ValueCell::ValueCell(const ValueCell& Number): numericValue(Number.numericValue), textValue(Number.textValue), isnumericValue(Number.isnumericValue){}

double ValueCell::getNumericValue() const{
    if(!isnumericValue){
        throw runtime_error("Ячейка содержит текстовое значение, числовое значение отсуствует");
    }
    return numericValue;
}

string ValueCell::getTextValue() const{
    if(isnumericValue){
        throw runtime_error("Ячейка содержит числовое значение, текстовое значение отсуствует");
    }
    return textValue;
}

void ValueCell::setNumericValue( double Numeric){
    numericValue = Numeric;
    isnumericValue = true;
}

void ValueCell::setTextValue( string Text){
    textValue = Text;
    isnumericValue = false;
}

bool ValueCell::isNumeric() const{
    return textValue.empty();
}

//-----------------------------------------------------------------------------------------------------------------

FormulaCell::FormulaCell(int startRow, int startCol, int endRow, int endCol, const vector<vector<ValueCell>>& cells)
    : startRow(startRow), startCol(startCol), endRow(endRow), endCol(endCol), cells(cells) {}

double FormulaCell::sum() const {
    double total = 0;
    for (int i = startRow; i <= endRow; ++i) {
        for (int j = startCol; j <= endCol; ++j) {
            if (cells[i][j].isNumeric()){

            total += cells[i][j].getNumericValue(); // Суммируем значения
        }
            else{
                throw runtime_error("Все значения ячеек должны быть с числовым типом данных");
            }}
    }
    return total;
}

double FormulaCell::product() const {
    double total = 1;
    for (int i = startRow; i <= endRow; ++i) {
        for (int j = startCol; j <= endCol; ++j) {
            if (cells[i][j].isNumeric()){
            total *= cells[i][j].getNumericValue(); // Перемножаем значения
            }
            else{
                throw runtime_error("Все значения ячеек должны быть с числовым типом данных");
            }
        }
    }
    return total;
}

double FormulaCell::average() const {
    double total = sum();
    int count = (endRow - startRow + 1) * (endCol - startCol + 1); // Подсчитываем количество ячеек
    return count > 0 ? total / count : 0;
}

void FormulaCell::printRange() const {
    cout << "Диапазон: (" << startRow << ", " << startCol << ") до (" 
            << endRow << ", " << endCol << ")" << endl;
}

void FormulaCell::printMatrix() const {
    cout << "Матрица значений:" << endl;
    for (int i = 0; i < cells.size(); ++i) {
        for (int j = 0; j < cells[i].size(); ++j) {
            if (cells[i][j].isNumeric()){
            cout << cells[i][j].getNumericValue() << " ";
        }else{
            cout << cells[i][j].getTextValue() << " ";
        }}
        cout << endl;
    }
}