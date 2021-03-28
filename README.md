# Python-Trading-Cards
A python program that allows players to maintain a deck of cards.

Trading Cards Game
Tray Ding Trading Co. have run a successful line of trading cards for many years; their “shiny” cards from an original production run now fetch upwards of £100 each on many online auction sites. Players take the management of their collection (colloquially known as their “deck”) of cards very seriously.

Tray Ding Trading now wish to release an app which allows players to enter the information about the cards in their deck, and for various statistics to be computed.

You have been approached to create a prototype of this in Python, to demonstrate the proof of concept.

Brief
You should design a python program that allows players to maintain a deck of cards. Each card should contain the following information:

Name

Type (one out of a given taxonomy of types)

Maximum Health Points

Moves (each move has a name and a damage factor, which is a whole number (0 or greater) that is deducted from an opposing card’s Health Points) Each card has between 1 and 5 moves.

Shiny status

Each creature represented on a trading card has a type. These are:

“Magi”

“Water”

“Fire”

“Earth”

“Air”

“Astral”


When managing their Deck, the player should be able to perform the following functions:

Add a card

Remove a card

View a card’s information

Obtain the most powerful card in their deck (the card in their deck with the highest average damage factor over all the card’s moves)

View the information of all shiny cards

View the information of all cards of a particular type

Technical specification
The following points must be incorporated into your program. If you do not incorporate the following, then the correctness of your program may not be able to be determined and you will lose marks. You may also lose marks for not following the brief.

Your submission is to be in a single file, which should be named TradingCards.py

You are to create the following object classes in Python:

Card

Deck

Within Card, you should have at least the following methods:

__init__(self, theName, theType, theHP, theMoves, isShiny)
This will initialise the Card instance

__str__(self)

You should use this to provide a sensible string representation of the card’s information, which will be used when viewing the card information

 

Within Deck, you should have at least the following methods (you may have helper methods to help you realise the functionality, but the functionality should come only from calls to these methods):

__init__(self)
This will initialise the Deck instance with the empty variables that are required.

inputFromFile(self, fileName)

This will populate the empty initialised Deck with the cards that are present in the .xlsx file <fileName>.xlsx. A sample file is included with this brief, so that you can test your program and learn the layout of the file. You are to use the openpyxl module to read your file.

__str__(self)
This should provide a sensible string representation of the Deck. Specifically, it should provide the total number of cards, the total number of shiny cards, and the average Damage value over the entire deck.

addCard(self, theCard)
This will add theCard to the deck

rmCard(self, theCard)
This will remove theCard from the deck

getMostPowerful(self)
This will return the Card that is most powerful (as described above)

getAverageDamage(self)
This will calculate and return the average damage inflicted by all cards in the deck. For example, say you have 3 cards in the deck, with average damages of 100, 150, 160. The average damage for your deck will be 136.6. This value should be displayed to 1 decimal point.

viewAllCards(self)
This will print the information of all the cards in the deck in a sensible way.

viewAllShinyCards(self)

This will print the information of all the shiny cards in the deck in a sensible way.

viewAllByType(self, theType)
This will print the information of all the cards in the deck that belong to the type of theType in a sensible way. The strings that are used to identify the different types are given above.

getCards(self)
This should return all cards held within the deck as a collection.

saveToFile(self, fileName)

This saves the Deck to an xlsx file that is called <fileName>.xlsx
It is to use the same format as the sample input file that is provided with this brief. You are to use the openpyxl module to read your file.

In the main body of this same file (ie: outside of the classes) you may have a method called main(), where you can test your classes. This is not compulsory, as both automatic unit testing (ie: that is run automatically) and the importing of your module into a Test Bed and manual verification will be used to determine the correctness of your program.

None of your code should run if the module has been imported (this is covered in Video Set 7). That is, all code should be encapsulated within methods and you should have something similar to the following code present, if it is needed:

if __name__ == “__main__”:

      main()

It is vitally important that you follow the brief exactly, especially when it comes to the class names and the method signatures. 

You should be very considerate to the design choices that you make; think very carefully about the structure of your object classes, the data types that you use in constructing your object classes, and the algorithms used to perform the various functions that are asked for. You should also carefully consider how you can make your program as robust as possible - try to use the most suitable tools that you know of, which have been introduced throughout the COMP517 material.

You should not import any libraries or modules into your code except for openpyxl. 
