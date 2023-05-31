#include "include_cpp/libxl.h"
#include <iostream>

using namespace libxl;

int main() 
{
   Book* book = xlCreateBook();

   if(book->load("demo.xls"))
   {
      Sheet* sheet = book->getSheet(0);
      if(sheet)
      {
         for(int row = sheet->firstRow(); row < sheet->lastRow(); ++row)
         {
            for(int col = sheet->firstCol(); col < sheet->lastCol(); ++col)
            {
               CellType cellType = sheet->cellType(row, col);
               std::wcout << "(" << row << ", " << col << ") = ";
               if(sheet->isFormula(row, col))
               {
                  const wchar_t* s = sheet->readFormula(row, col);
                  std::wcout << (s ? s : L"null") << " [formula]";
               }
               else
               {
                  switch(cellType)
                  {
                     case CELLTYPE_EMPTY: std::wcout << "[empty]"; break;
                     case CELLTYPE_NUMBER:
                     {
                        double d = sheet->readNum(row, col);
                        std::wcout << d << " [number]";
                        break;
                     }
                     case CELLTYPE_STRING:
                     {
                        const wchar_t* s = sheet->readStr(row, col);
                        std::wcout << (s ? s : L"null") << " [string]";
                        break;        
                     }
                     case CELLTYPE_BOOLEAN:
                     {
                        bool b = sheet->readBool(row, col);
                        std::wcout << (b ? "true" : "false") << " [boolean]";
                        break;
                     }
                     case CELLTYPE_BLANK: std::wcout << "[blank]"; break;
                     case CELLTYPE_ERROR: std::wcout << "[error]"; break;
                  }
               }
               std::wcout << std::endl;
            }
         }
      }
   }

   book->release();
   return 0;
}