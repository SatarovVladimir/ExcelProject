
#ifndef LAB2_EXCEL_H
#define LAB2_EXCEL_H
#include <iostream>
#include <cassert>
#include <iostream>
#include <vector>
#include <stdexcept>

using namespace std;

using namespace std;

class ValueCell{
private:
    double numericValue;
    string textValue;
    bool isnumericValue;
public:
    ValueCell();
    ValueCell(double Numeric);
    ValueCell(string Text);
    ValueCell(const ValueCell& Number);

    double getNumericValue() const;
    string getTextValue() const;
    void setNumericValue( double Numeric);
    void setTextValue( string Text);
    bool isNumeric() const;
};

class FormulaCell : public ValueCell{
private:
    vector<vector<ValueCell>> cells;
    int startRow, startCol, endRow, endCol;
    int beginnigRange, endingRange;
public:
    FormulaCell(int startRow, int startCol, int endRow, int endCol, const vector<vector<ValueCell>>& cells);

    double sum() const;
    double product() const;
    double average() const;

    void printRange() const;
    void printMatrix() const;
};


#endif //LAB2_EXCEL_H
