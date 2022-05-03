#include <iostream>
#include "libxl.h"

using namespace libxl;
using namespace std;

int main(){
    cout << "실행" << endl;
    Book* book = xlCreateBook();

    if(book->load("test.xlsx")){
        cout << "파일 읽음" << endl;
        Sheet* sheet = book->getSheet(0);
        if(sheet){
            for(int row = sheet->firstRow(); row < sheet->lastRow(); row++){
                for(int col = sheet->firstCol(); col < sheet->lastCol(); col++){
                    CellType cellType = sheet->cellType(row, col);
                    std::wcout << "(" << row << ", " << col << ") = ";
                    /*
                    switch(cellType){
                        case CELLTYPE_EMPTY: std::wcout << "[empty]"; break;
                        case CELLTYPE_NUMBER:
                            double d = sheet->readNum(row, col);
                            std::wcout << d << " [number]";
                            break;
                        case CELLTYPE_STRING:
                            const char* s = sheet->readStr(row, col);
                            wcout << s << " [strings]";
                            break;
                        case CELLTYPE_BOOLEAN: break;
                        case CELLTYPE_BLANK: break;
                        case CELLTYPE_ERROR: std::wcout << "[error]"; break;                        
                    }
                    */
                    wcout << endl;
                }
            }
        }
    }
    

    book->release();
    return 0;
}