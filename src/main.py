from services.parser import save_excel


def main():
    try:
        limit = int(input('How many cards do you want? Enter a number.\nThe program may take a long time to run when you enter large numbers.\n'))
        if limit < 1:
            print("Incorrect value, try again. The value must be greater than 0")
            return
    except Exception as e:
        print(f"Error: {e}. Try again.")
        return
    save_excel(limit=limit)


if __name__ == "__main__":
    print('Programm start...\n')
    main()
    print('The program has completed its work\n')
