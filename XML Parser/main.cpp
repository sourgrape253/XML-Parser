//
//	Bachelor of Software Engineering
//	Media Design School
//	Auckland
//	New Zealand
//
//	(c) 2005 - 2015 Media Design School
//
//	File Name	:	main.cpp
//	Description	:	main source file containing the XML parser
//	Author		:	Chris Stone
//	Mail		:	christopher.sto6279@mediadesignschool.com
//

// PugiXML include
#include "PugiXML\pugixml.hpp"
#include "libxl.h"

// Library includes
#include <Windows.h>
#include <iostream>
#include <string>
#include <vector>

// pugi and libxl namespaces
using namespace pugi;
using namespace libxl;

/*
*	Opens an XML file and uses the information to form an Excel sheet
*	Parameters:	none
*	Retrurns: void
*/
void XMLtoXLS()
{
	std::cout << "Parsing...";

	// Create and Excel spread sheet
	Book* book = xlCreateBook();
	if (!book) return;

	// Create a new sheet
	Sheet* sheet = book->addSheet("Sheet1");
	if (!sheet) return;

	// Setup the formats for headings
	Font* font = book->addFont();
	font->setBold(true);
	Format* boldFormat = book->addFormat();
	boldFormat->setFont(font);

	// Load the xml document from file
	xml_document doc;
	xml_parse_result result = doc.load_file("XML Folder\\Games.xml");
	xml_node favouriteGames = doc.child("FavouriteGames");

	// Write the headings from the xml to the excel sheet
	int x = 1;
	for (xml_node headings : favouriteGames.first_child().children())
	{
		sheet->writeStr(1, x, headings.name(), boldFormat);
		++x;
	}

	// Fill in the information
	x = 0;
	int y = 2;
	for (xml_node game : favouriteGames.children("Game"))
	{
		sheet->writeStr(y, x, game.attribute("name").value(), boldFormat);
		for (xml_node values : game.children())
		{
			++x;
			sheet->writeStr(y, x, values.first_child().value());
		}
		++y;
		x = 0;
	}

	system("CLS");
	std::cout << "Opening...";

	// Save the excel sheet and open it
	if (book->save("FavouriteGames.xls"))
		::ShellExecute(NULL, "open", "FavouriteGames.xls", NULL, NULL, SW_SHOW);
	else
		std::cout << book->errorMessage() << std::endl;

	// Release the memory occupied by the book
	book->release();
	system("CLS");
}

/*
*	Entry point of the application. Runs the interface loop
*	Parameters:	none
*	Retrurns: int
*/
int main()
{
	while (true)
	{
		std::string userChoice = "0";
		bool validAnswer = true;
		do
		{
			std::cout << "Would you like to:\n\n  1)\tParse XML to XLS\n  2)\tQuit\n\n";

			// Have the user enter their answer
			if (!validAnswer)
				std::cout << " Invalid Answer!\n";
			std::cout << "Enter (1), (2):  ";
			std::getline(std::cin, userChoice);

			// Check answer is 1 or 2
			if (userChoice != "1" && userChoice != "2")
				validAnswer = false;
			else
				validAnswer = true;

			// Clear the screen
			system("CLS");
		} while (validAnswer == false);

		if (userChoice == "1")
			XMLtoXLS();
		if (userChoice == "2")
			break;
	}

	return 0;
}