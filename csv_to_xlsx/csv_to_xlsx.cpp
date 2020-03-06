
#include <iostream>
#include <fstream>
#include <string>
#include <sstream>
#include <xlnt/xlnt.hpp>
#include <regex>

using namespace std;
using namespace xlnt;

string ReplaceAll(std::string str, const std::string& from, const std::string& to) {
	size_t start_pos = 0;
	while ((start_pos = str.find(from, start_pos)) != std::string::npos) {
		str.replace(start_pos, from.length(), to);
		start_pos += to.length(); // Handles case where 'to' is a substring of 'from'
	}
	return str;
}

//replaces one char by two, used for ANSI to UTF-8, needs proper rework
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

//Handle UTF-8 Cases, needs proper rework
string to_utf8(string line)
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
	string line;
	stringstream ss;

	//Handle if any File is present
	if (argv[1] == NULL)
	{
		cout << "No File at all" << endl;
		return 0;
	}
	for (int i = 1; i < argc; i++)
	{
		//INIT Variables
		string filename = argv[i];
		input.open(filename, ios::in);
		workbook wb;
		worksheet ws = wb.active_sheet();

		int row = 1;
		int column = 1;

		while (getline(input, line, '\n'))
		{
			column = 1;

			//move string into stringstream so that getline will work
			ss << line;
			while (getline(ss, line, ';'))
			{
				//Write Data, check Data type
				if (regex_match(line, regex("\\d+,\\d+")) or regex_match(line, regex("\\d+\\.\\d+")))
				{
					line = ReplaceAll(line, ",", ".");
					//matched double
					double value = stod(line);
					ws.cell(column, row).value(value);
				}
				else if(regex_match(line, regex("\\d+")))
				{
					//matched integer
					int value = stoi(line);
					ws.cell(column, row).value(value);
				}
				else
				{
					ws.cell(column, row).value(to_utf8(line));
				}				
				column++;
			}

			//clear stringstream so getline will start over
			ss.clear();
			row++;
		}

		//save xlsx
		wb.save(filename.substr(0, filename.find_last_of(".")) + ".xlsx");
		input.close();
	}
}
