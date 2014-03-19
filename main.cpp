#include <iostream>
#include <stdio.h>
#include <stdlib.h>
#include <fstream>
#include <boost/algorithm/string.hpp>
#include <string.h>
#include <iomanip>
#include <myiconv/iconvlite.h>

#include <Xlsx/Workbook.h>

using namespace std;
using namespace SimpleXlsx;

double get_kurs();
string myReplace(string data,string needle,string payload);

int main()
{
    const double Kurs = get_kurs();
    cout << "[+] Kurs: " << Kurs << endl;
    ifstream InFile ("in2.txt");
    vector<string> Data;
    vector<string> cells;
    vector<string> outcells;
    string line;
    long c=0;
    int cols;
    int r;
    double Num;
    string tmp;

    CWorkbook book;
    CWorksheet &sheet = book.AddSheet(_T("1"));
    string utf;
    vector<CellDataStr> data;   // (data:style_index)
    CellDataStr def;
    def.style_id = 0;

    if (InFile.is_open()){
        while ( InFile.good() ){
            getline (InFile,line);
            Data.push_back(line);
            //boost::trim(Data[0]);
            boost::split(cells, Data[0], boost::is_any_of("\t"));
            if(c==0){
            //first line
                cells[8] = utf2cp("Кроссы ОЕ");
                cells[4] = utf2cp("ЦенаЕвро");
                cells.push_back(utf2cp("ЦенаРубли"));
                cols = cells.size();
                c++;
            } else {
            //main case
                r=14;
                tmp = cells[r];
                std::size_t found = tmp.find("_");
                if (found!=std::string::npos) {
                    tmp.replace(found,tmp.length()-found,"");
                }
                cells[r] = tmp;

                r=4;
                tmp = cells[r];
                std::replace( tmp.begin(), tmp.end(), ',', '.');
                Num = ::atof(tmp.c_str());
                Num = Num*Kurs;
                std::ostringstream buffer;
                buffer <<setprecision(10)<< Num;
                tmp = buffer.str();
                std::replace( tmp.begin(), tmp.end(), '.', ',');
                cells.push_back(tmp);

                r=5;
                tmp = cells[r];
                Num = ::atof(tmp.c_str());
                if(Num>16) {
                    Num=16;
                    std::ostringstream buffer;
                    buffer << Num;
                    cells[r] = buffer.str();
                }
            }
            outcells.push_back(cells[0]);
            outcells.push_back(cells[14]);
            outcells.push_back(cells[1]);
            outcells.push_back(cells[3]);
            outcells.push_back(cells[2]);
            outcells.push_back(cells[8]);
            outcells.push_back(cells[5]);
            outcells.push_back(cells[4]);
            outcells.push_back(cells[15]);
            outcells.push_back(cells[13]);

            for (int j = 0; j < 10; j++) {
                utf = outcells[j];
                utf = myReplace(utf,"&","&amp;");
                utf = myReplace(utf,"<","&lt;");
                utf = myReplace(utf,">","&gt;");
                def.value = cp2utf(utf);
                data.push_back(def);
            }
            sheet.AddRow(data);
            data.clear();
            Data.clear();
            cells.clear();
            outcells.clear();
        }
        InFile.close();
        bool bRes = book.Save(_T("out.xlsx"));
        if (bRes)   cout << "[+] The xlsx has been saved successfully" << endl;
        else        cout << "[-] The xlsx saving has been failed" << endl;
    } else {
        cout << "[-] Cant load file" << endl;
        return -1;
    }
    return 0;
}

string myReplace(string data,string needle,string payload){
    std::size_t found = data.find(needle);
    if (found!=std::string::npos) {
        boost::replace_all(data, needle, payload);
    }
    return data;
}

double get_kurs(){
    string sKurs;
    cout << "\nKurs: ";
    getline(cin,sKurs);
    if (!sKurs.empty())
    {
        std::replace( sKurs.begin(), sKurs.end(), ',', '.');
        return ::atof(sKurs.c_str());
    }
    else
        return 44.5;
}
