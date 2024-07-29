# Luxor Hotel Management System

![Hotel Luxor](https://github.com/JoacoMarc/Hotel-Guest-Management-System/blob/main/HotelLuxorBanner.jpg)

Welcome to the Luxor Hotel Management System! This program is designed to manage hotel guests, including check-ins, room assignments, and various guest-related queries.

[![Watch the video](https://github.com/JoacoMarc/Hotel-Guest-Management-System/blob/main/LuxorPreview.png)](https://drive.google.com/file/d/1qzy2rXzdfiMitIJUhM-W6MUScAHX8TpF/view?usp=drive_link)

## 📋 Table of Contents
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Examples](#examples)
- [Contributing](#contributing)
- [License](#license)

## 🌟 Features
- ➕ **Add new guests** with their details
- 🏨 **Assign guests** to specific rooms
- 🔍 **Find guest information** by last name
- 📈 **View the most occupied floor**
- 🏢 **View the number of empty rooms**
- 👥 **View the floor with the most occupants**
- 🚪 **View rooms with upcoming departures**
- 🖼️ **Visual representation of the hotel's occupancy**
- 📊 **Create an Excel file with guest information**
- 🧹 **Clear the guest file and vacate the hotel**

## 🛠️ Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/hotel-luxor-management.git
   cd hotel-luxor-management
   ```

2. Create and activate a virtual environment:
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`
   ```

3. Install the dependencies:
   ```bash
   pip install -r requirements.txt
   ```

### Requirements
Ensure you have the following libraries installed:
- `tkinter` (included with Python)
- `Pillow`
- `openpyxl`
- `colorama`

Install these libraries using pip if necessary:
```bash
pip install pillow openpyxl colorama
```

## 🚀 Usage
Run the main program:
```bash
python main.py
```

## 📝 Examples
### 🏠 Main Menu
Upon running the program, you will be greeted with the following menu:

```
Menu:
1. Enter new guest
2. Find guest by last name
3. Most occupied floor
4. Number of empty rooms
5. Floor with most occupants
6. Next departures
7. View hotel
8. Create or update guest Excel
9. Clear guest Excel and vacate hotel
10. Exit

Choose an option:
```

### ➕ Adding a New Guest
- Select option `1` to add a new guest.
- Follow the prompts to enter guest details and assign a room.

### 🏨 Viewing the Hotel
- Select option `7` to view the current occupancy of the hotel. The output will display a graphical representation of the hotel layout with color-coded rooms.

### 🔍 Finding a Guest by Last Name
- Select option `2` and enter the last name to search for a guest. The program will display the guest's information and room number.

### 📊 Creating a Guest File
- Select option `8` to create an Excel file (`guests.xlsx`) with all guest information sorted by last name. Each occupant will have their own row to ensure the data is properly aligned.

### 🧹 Clearing the Guest File and Vacating the Hotel

- Select option `9` to clear the guest file and vacate the hotel. This will reset all guest information and create a new empty Excel file.

## 🛠️ Code Overview

- **Card Hierarchy**: Defined using a dictionary `jerarquia` with cards as keys and their ranks as values.
- **Card Distribution**: The function `distribuircartas` handles the random distribution of cards to players.
- **Display Cards**: The function `mostrar_cartas` prints each player's hand.
- **Player Choice**: The function `elegir_carta` allows players to choose a card to play.
- **Play Rounds**: 
  - The function `jugar_ronda` manages the logic for playing a single round and determining the winner.
  - The function `jugar_desempate` handles tie-breaker rounds.
- **Game Loop**: The main loop runs through the rounds, checks for winners, and handles the game flow.

## 🤝 Contributing
Contributions are welcome! Please fork this repository and submit a pull request for any features, bug fixes, or improvements.

## 📜 License
This project is licensed under the MIT License - see the LICENSE file for details.
```
