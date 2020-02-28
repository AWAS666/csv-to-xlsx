
#include <iostream>
#include <fstream>
#include <string>
#include <sstream>
#include <xlnt/xlnt.hpp>

using namespace std;
using namespace xlnt;

//ersetze einzelnes Char in String mit String
void replace_char_by_str(string& input, char character, char char1,char char2)
{
	for (int i = 0; i < input.length(); i++)
	{
		string start;
		string ende;
		if (input[i] == character)
		{
			if (i > 0)
			{
				start = input.substr(0, i);
			}
			if (i < input.length() - 1)

			{
				ende = input.substr(i + 1, input.length() - i);
			}
			input = start + char1+ char2 + ende;
		}
	}
}

//Umlautsäuberung für UTF-8
string umlaut_utf8(string line)
{
	//https://www.i18nqa.com/debug/utf8-debug.html
	replace_char_by_str(line, 0xE4, 0xC3,0x80);
	replace_char_by_str(line, 0xFC, 0xC3,0xBC);
	replace_char_by_str(line, 0xF6, 0xC3,0xB6);
	replace_char_by_str(line, 0xC4, 0xC3 ,0x84);
	replace_char_by_str(line, 0xD6, 0xC3 ,0x96);
	replace_char_by_str(line, 0xDC, 0xC3 ,0x9C);
	return line;
}


int main(int argc, char* argv[])
{
	ifstream input;
	if (argv[1] == NULL)
	{
		cout << "Fehler bei Parametereingabe" << endl;
		return 0;
	}
	//INIT
	string filename = argv[1];
	input.open(filename, ios::in);
	workbook wb;
	worksheet ws = wb.active_sheet();
	
	string line;

	stringstream ss;
	int Reihe = 1;
	int Spalte = 1;

	while (getline(input, line, '\n'))
	{
		Spalte = 1;		
		ss << line;
		while (getline(ss, line, ';'))
		{			
			//Schreibe Zelle
			ws.cell(Spalte, Reihe).value(umlaut_utf8(line));
			Spalte++;			
		}
		ss.clear();
		Reihe++;
	}
	wb.save(filename.substr(0, filename.find_last_of("."))+ ".xlsx");

}
