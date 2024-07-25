import random
from datetime import datetime
from tabulate import tabulate
from colorama import Fore, Style, init
from openpyxl import Workbook

# Initialize colorama
init()

def generate_matrix(rows=10, columns=6):
    """Generates an empty matrix of size rows x columns."""
    return [[0 for _ in range(columns)] for _ in range(rows)]

def print_matrix(matrix):
    """Prints the matrix in a readable format."""
    display_matrix = [[room["occupants"] if room != 0 else 0 for room in floor] for floor in matrix]
    display_matrix.reverse()  # Reverse the matrix to print from top floor to bottom floor
    floor_numbers = [[f"Floor {10 - i}"] + display_matrix[i] for i in range(len(display_matrix))]
    headers = [" "] + [f"Room {i+1}" for i in range(len(display_matrix[0]))]
    print(tabulate(floor_numbers, headers, tablefmt="fancy_grid"))

def validate_dates(start_date, end_date):
    """Validates that the start date is before the end date."""
    start = datetime.strptime(str(start_date), '%d%m%Y')
    end = datetime.strptime(str(end_date), '%d%m%Y')
    return start < end

def assign_room(id, first_name, last_name, birth_date, check_in, check_out, occupants, matrix):
    """Assigns a specified room to a guest."""
    guest = {"id": id, "first_name": first_name, "last_name": last_name, 
             "birth_date": birth_date, "check_in": check_in, "check_out": check_out, 
             "occupants": occupants}

    # If there are more occupants, ask for their details
    if occupants > 1:
        additional_occupants = []
        for i in range(occupants - 1):
            print(Fore.CYAN + f"Enter details for occupant {i+2}:" + Style.RESET_ALL)
            occupant_id = int(input("  Enter ID: "))
            occupant_first_name = input("  Enter first name: ")
            occupant_last_name = input("  Enter last name: ")
            occupant_birth_date = int(input("  Enter birth date (ddmmyyyy): "))
            additional_occupants.append({"id": occupant_id, "first_name": occupant_first_name, 
                                         "last_name": occupant_last_name, "birth_date": occupant_birth_date})
        guest["additional_occupants"] = additional_occupants
    else:
        guest["additional_occupants"] = []

    while True:
        row = int(input(Fore.CYAN + "Enter floor (1-10): " + Style.RESET_ALL)) - 1
        col = int(input(Fore.CYAN + "Enter room number (1-6): " + Style.RESET_ALL)) - 1
        if matrix[row][col] == 0:
            matrix[row][col] = guest
            break
        else:
            print(Fore.RED + "Room already occupied, please choose another." + Style.RESET_ALL)

    return matrix

def find_guest(matrix, last_name):
    """Finds and returns the room and guest data for a given last name."""
    for floor_index, floor in enumerate(matrix):
        for room_index, room in enumerate(floor):
            if room != 0 and room["last_name"].lower() == last_name.lower():
                return (floor_index, room_index, room)
    return None

def print_guest_info(guest):
    """Prints guest information in a readable format."""
    print(Fore.GREEN + f"Guest Information:\n"
          f"  ID: {guest['id']}\n"
          f"  First Name: {guest['first_name']}\n"
          f"  Last Name: {guest['last_name']}\n"
          f"  Birth Date: {guest['birth_date']}\n"
          f"  Check-in Date: {guest['check_in']}\n"
          f"  Check-out Date: {guest['check_out']}\n"
          f"  Number of Occupants: {guest['occupants']}" + Style.RESET_ALL)
    
    if guest["additional_occupants"]:
        print(Fore.GREEN + "  Additional Occupants:" + Style.RESET_ALL)
        for i, occupant in enumerate(guest["additional_occupants"], start=1):
            print(Fore.GREEN + f"    Occupant {i}:\n"
                  f"      ID: {occupant['id']}\n"
                  f"      First Name: {occupant['first_name']}\n"
                  f"      Last Name: {occupant['last_name']}\n"
                  f"      Birth Date: {occupant['birth_date']}" + Style.RESET_ALL)

def most_occupied_floor(matrix):
    """Returns the most occupied floor."""
    occupancy = [sum(1 for room in floor if room != 0) for floor in matrix]
    most_occupied = occupancy.index(max(occupancy))
    return most_occupied + 1

def empty_rooms(matrix):
    """Counts the number of empty rooms."""
    return sum(1 for floor in matrix for room in floor if room == 0)

def floor_with_most_occupants(matrix):
    """Returns the floor with the most occupants."""
    occupants = [sum(room["occupants"] for room in floor if room != 0) for floor in matrix]
    return occupants.index(max(occupants)) + 1

def next_departures(matrix, current_date):
    """Finds the rooms with the closest upcoming departures."""
    current = datetime.strptime(str(current_date), '%d%m%Y')
    departures = [(floor, room, datetime.strptime(str(room["check_out"]), '%d%m%Y')) 
                  for floor in matrix for room in floor if room != 0]

    days_diff = [(floor, room, (departure - current).days) 
                 for floor, room, departure in departures]

    closest_days = min(days_diff, key=lambda x: x[2])[2]
    next_rooms = [room for floor, room, days in days_diff if days == closest_days]

    return next_rooms

def sort_by_last_name(guest):
    """Sort function by last name."""
    return guest["last_name"]

def create_guest_file(matrix):
    """Creates an Excel file with guest information."""
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Guests"

        headers = ["ID", "First Name", "Last Name", "Birth Date", "Check-in Date", "Check-out Date", "Occupants"]

        ws.append(headers)

        guests = [room for floor in matrix for room in floor if room != 0]
        guests.sort(key=sort_by_last_name)

        for guest in guests:
            guest_list = [guest["id"], guest["first_name"], guest["last_name"], guest["birth_date"], 
                          guest["check_in"], guest["check_out"], guest["occupants"]]
            ws.append(guest_list)
            for occupant in guest["additional_occupants"]:
                occupant_list = ["", "", "", "", "", "", ""]
                occupant_list.extend([occupant["id"], occupant["first_name"], occupant["last_name"], 
                                      occupant["birth_date"]])
                ws.append(occupant_list)

        wb.save("guests.xlsx")
        print(Fore.GREEN + "Guest file created successfully." + Style.RESET_ALL)

    except FileNotFoundError:
        print(Fore.RED + "File not found." + Style.RESET_ALL)
    except OSError as msg:
        print(Fore.RED + "Error:", msg, Style.RESET_ALL)

def clear_guest_file():
    """Clears the Excel file by creating a new empty file."""
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Guests"

        headers = ["ID", "First Name", "Last Name", "Birth Date", "Check-in Date", "Check-out Date", "Occupants"]
        ws.append(headers)

        wb.save("guests.xlsx")
        print(Fore.GREEN + "Guest file cleared successfully." + Style.RESET_ALL)

    except FileNotFoundError:
        print(Fore.RED + "File not found." + Style.RESET_ALL)
    except OSError as msg:
        print(Fore.RED + "Error:", msg, Style.RESET_ALL)

def main():
    """Main function to handle the program logic."""
    print(Fore.MAGENTA + Style.BRIGHT + "=======================")
    print(Fore.CYAN + Style.BRIGHT + "WELCOME TO LUXOR HOTEL")
    print(Fore.MAGENTA + Style.BRIGHT + "=======================", end="\n\n")
    
    id_list = []
    matrix = generate_matrix()

    while True:
        print(Fore.YELLOW + "\nMenu:\n1. Enter new guest\n2. Find guest by last name\n3. Most occupied floor\n4. Number of empty rooms\n5. Floor with most occupants\n6. Next departures\n7. View hotel\n8. Create or update guest Excel\n9. Clear guest Excel and vacate hotel\n10. Exit" + Style.RESET_ALL)
        choice = input(Fore.CYAN + "Choose an option: " + Style.RESET_ALL)
        
        if choice == '1':
            try:
                id = int(input(Fore.CYAN + "Enter ID (-1 to exit): " + Style.RESET_ALL))
                if id == -1:
                    continue
                assert len(str(id)) == 8, Fore.RED + "ID must have 8 digits." + Style.RESET_ALL
                assert id not in id_list, Fore.RED + "ID already exists, please enter another." + Style.RESET_ALL
                id_list.append(id)

                first_name = input(Fore.CYAN + "Enter first name: " + Style.RESET_ALL)
                last_name = input(Fore.CYAN + "Enter last name: " + Style.RESET_ALL)
                birth_date = int(input(Fore.CYAN + "Enter birth date (ddmmyyyy): " + Style.RESET_ALL))
                check_in = int(input(Fore.CYAN + "Enter check-in date (ddmmyyyy): " + Style.RESET_ALL))
                check_out = int(input(Fore.CYAN + "Enter check-out date (ddmmyyyy): " + Style.RESET_ALL))
                assert validate_dates(check_in, check_out), Fore.RED + "Check-in date must be before check-out date." + Style.RESET_ALL
                    
                occupants = int(input(Fore.CYAN + "Enter number of occupants: " + Style.RESET_ALL))

                matrix = assign_room(id, first_name, last_name, birth_date, check_in, check_out, occupants, matrix)

            except ValueError:
                print(Fore.RED + "Please enter valid numbers for ID, birth date, check-in, and check-out dates." + Style.RESET_ALL)
            except AssertionError as msg:
                print(Fore.RED + "Error: " + str(msg) + Style.RESET_ALL)
                if id in id_list:
                    id_list.remove(id)

        elif choice == '2':
            last_name = input(Fore.CYAN + "Enter last name to search: " + Style.RESET_ALL)
            result = find_guest(matrix, last_name)
            if result:
                floor, room, guest = result
                print(Fore.GREEN + f"Guest found in floor {floor + 1}, room {room + 1}:" + Style.RESET_ALL)
                print_guest_info(guest)
            else:
                print(Fore.RED + "Guest not found." + Style.RESET_ALL)
        
        elif choice == '3':
            print(Fore.GREEN + "\nMost occupied floor is:", most_occupied_floor(matrix), Style.RESET_ALL)

        elif choice == '4':
            print(Fore.GREEN + "\nNumber of empty rooms:", empty_rooms(matrix), Style.RESET_ALL)

        elif choice == '5':
            print(Fore.GREEN + "\nFloor with most occupants:", floor_with_most_occupants(matrix), Style.RESET_ALL)

        elif choice == '6':
            current_date = int(input(Fore.CYAN + "Enter current date (ddmmyyyy): " + Style.RESET_ALL))
            print(Fore.GREEN + "\nRooms with upcoming departures:\n" + Style.RESET_ALL)
            upcoming_rooms = next_departures(matrix, current_date)
            for room in upcoming_rooms:
                print(Fore.YELLOW + "Room of:", room["last_name"] + Style.RESET_ALL)

        elif choice == '7':
            print_matrix(matrix)

        elif choice == '8':
            create_guest_file(matrix)

        elif choice == '9':
            id_list.clear()
            matrix = generate_matrix()
            clear_guest_file()

        elif choice == '10':
            break

if __name__ == "__main__":
    main()


