import os
import random
import getpass

# List of quotes
QUOTES = [
    '"You have no one to blame but yourself for who you are"',
    '"You can\'t escape the darkness within you, for it is your own creation."',
    '"Embrace your demons, for they are the shadows that define your light."',
    '"In the end, the only chains that bind you are the ones you forged yourself."',
    '"The mirror reflects your truth, whether you choose to see it or not."',
    '"Your destiny is written by your own hand, in ink as dark as your soul."',
    '"You are the architect of your own fate, sculpting it with every choice you make."',
    '"The scars on your heart are not inflicted by others, but by your own hand."',
    '"You are the master of your own universe, painting it with the colors of your desires."',
    '"To blame fate is to deny the power you hold over your own life."',
    '"The greatest battles are fought within, where the demons of doubt and fear reside."',
    '"You are both the hero and the villain in your own story, playing each role with equal fervor."',
    '"The echoes of your choices reverberate through eternity, shaping the very fabric of your existence."',
    '"You are not a prisoner of circumstance, but a slave to your own convictions."',
    '"The road to redemption is paved with the stones of accountability."',
    '"In the theater of life, you are both the actor and the audience, applauding your triumphs and mourning your tragedies."',
    '"Your past mistakes are the stepping stones to enlightenment, leading you towards the path of self-discovery."',
    '"To know oneself is to conquer the greatest adversary – the self."',
    '"The shadows of your past do not define you, unless you choose to let them."',
    '"You are the sum of your choices, the culmination of every decision you\'ve ever made."',
    '"In the end, the only judgement that truly matters is the one you pass upon yourself."'
]

def display_quote():
    # Print random quote
    print(random.choice(QUOTES))
    # Print signature line for Lucifer Morningstar
    print("\n                                                - Lucifer Morningstar\n\n")

def set_git_credentials():
    attempts = 3
    while attempts > 0:
        try:
            print("\nSetting up Git credentials...\n")
            # Ask for username and password
            username = input("Enter your Git username: ")
            password = getpass.getpass("Enter your Git password: ")
            
            # Configure Git to use credential helper store
            os.system("git config --global credential.helper store")
            # Get the full path to the home directory of the current user
            home_dir = os.path.expanduser("~")
            # Store the credentials in the home directory
            credentials_file = os.path.join(home_dir, ".git-credentials")
            with open(credentials_file, "w") as f:
                f.write(f"https://{username}:{password}@gitlab.kpit.com/ketakic/DRT_Jenkins.git\n")
            print("\nGit credentials configured and saved successfully.\n")
            return True
        except Exception as e:
            attempts -= 1
            print(f"\nAn error occurred while setting up Git credentials: {e}. You have {attempts} attempts left.\n")

    if attempts == 0:
        print("\nYou have exceeded the maximum number of attempts. Please try again later.\n")
        return False
        
def remove_git_credentials():
    try:
        print("Removing Git credentials...\n")
        # Remove the stored credentials
        credentials_file = os.path.expanduser("~/.git-credentials")
        if os.path.exists(credentials_file):
            os.remove(credentials_file)
            print("Git credentials removed successfully.")
        else:
            print("Git credentials file not found.")
    except Exception as e:
        print("\nAn error occurred while removing Git credentials:\n", e)

def clone_repository():
    try:
        print("\nChecking if repository exists...\n")
        # Check if the repository is already cloned
        if not os.path.exists("D:/Git_Data/DRT_Jenkins/.git"):
            print("Repository not found. Cloning...\n")
            # Clone the repository if it's not already cloned
            os.makedirs("D:/Git_Data", exist_ok=True)
            os.chdir("D:/Git_Data")
            os.system("git clone https://gitlab.kpit.com/ketakic/DRT_Jenkins.git")
            print("Repository cloned successfully.")
        else:
            print("Repository already exists. Skipping cloning.\n")
    except Exception as e:
        print("\nAn error occurred during repository cloning:\n", e)

def switch_pull():
    try:
        print("\nChanging directory to the cloned repository...\n")
        # Change directory to the cloned repository
        os.chdir("D:/Git_Data/DRT_Jenkins")
        print("Switching to the 'drt' branch and pulling latest changes...\n")
        # Switch to the 'drt' branch and pull latest changes
        os.system("git checkout drt")
        os.system("git pull origin drt")
        print("\n'drt' branch switched and latest changes pulled successfully.\n")
    except Exception as e:
        print("\nAn error occurred during switching/pulling:\n", e)

def add_commit_push():
    try:
        print("Changing directory to the cloned repository...\n")
        # Change directory to the cloned repository
        os.chdir("D:/Git_Data/DRT_Jenkins")
        print("Adding changes to the staging area...\n")
        # Add all changes
        os.system("git add .")
        print("Committing changes...\n")
        # Prompt the user to enter commit message
        commit_message = input("Enter commit message: ")
        if commit_message:
            # Commit changes with the entered message
            os.system(f'git commit -m "{commit_message}"')
            print("Pushing changes to the 'drt' branch...\n")
            # Push changes to the 'drt' branch
            os.system("git push origin drt")
            print("\nChanges added, committed, and pushed successfully to 'drt' branch.\n")
        else:
            print("Commit message cannot be empty. No changes made.\n")
    except Exception as e:
        print("\nAn error occurred during adding/committing/pushing:\n", e)

def merge_main():
    try:
        print("\nChanging directory to the cloned repository...\n")
        # Change directory to the cloned repository
        os.chdir("D:/Git_Data/DRT_Jenkins")
        print("Merging changes from 'main' into 'drt' branch...\n")
        # Merge changes from 'main' into 'drt' branch
        os.system("git checkout main")
        os.system("git pull origin main")
        os.system("git checkout drt")
        os.system("git merge main")
        print("\nChanges merged successfully from 'main' into 'drt' branch.\n")
    except Exception as e:
        print("\nAn error occurred during merging:\n", e)

def lock_unlock_file():
    try:
        file_path = "D:/Git_Data/DRT_Jenkins/Test_data/GoldenData/Variable/GlobalVariableDataV1.0.xlsx"
        print("Changing directory to the cloned repository...\n")
        # Change directory to the cloned repository
        os.chdir("D:/Git_Data/DRT_Jenkins")
        print("Pulling latest changes from 'drt' branch...\n")
        # Pull latest changes from 'drt' branch
        os.system("git pull origin drt")
        print("Locking the Global Variable Sheet...\n")
        # Lock the Global Variable Sheet
        os.system(f"git lfs lock {file_path}")
        print("Global Variable Sheet locked successfully.\n")
        print("Make changes in the global variable sheet.\n")
        # Ask the user if changes are made
        changes_made = input("Have you made the necessary changes? (Yes/No): ").lower()
        if changes_made == "yes":
            print("Pushing the updated data to GitLab...\n")
            # Push changes to the 'drt' branch
            os.system("git push origin drt")
            print("Data pushed successfully to GitLab.\n")
            print("Unlocking the Global Variable Sheet...\n")
            # Unlock the Global Variable Sheet
            os.system(f"git lfs unlock {file_path}")
            print("Global Variable Sheet unlocked successfully.")
        else:
            print("Data not pushed. Aborting locking and unlocking.")
    except Exception as e:
        print("An error occurred during locking/unlocking:", e)

if __name__ == "__main__":
    try:
        display_quote()
        print("\nWelcome to the Git Operations Script\n")

        # Check if Git credentials are already configured
        if not os.path.exists(os.path.expanduser("~/.git-credentials")):
            set_credentials_choice = input("Do you want to set up Git credentials now? (Yes/No): ").lower()
            if set_credentials_choice == "yes":
                set_git_credentials()

        while True:
            try:
                print("\nWhat operation would you like to perform?\n")
                print("1. Clone or update the repository")
                print("2. Switch to the 'drt' branch and pull latest changes")
                print("3. Add, commit, and push changes to 'drt' branch")
                print("4. Merge changes from 'main' into 'drt' branch")
                print("5. Lock and unlock the Global Variable Sheet")
                print("6. Exit")
                choice = input("\nEnter your choice (1-6): ")

                if choice == "1":
                    clone_repository()
                elif choice == "2":
                    switch_pull()
                elif choice == "3":
                    add_commit_push()
                elif choice == "4":
                    merge_main()
                elif choice == "5":
                    lock_unlock_file()
                elif choice == "6":
                    break
                else:
                    print("Invalid choice. Please enter a number between 1 and 6.")
            except KeyboardInterrupt:
                print("\nExiting the script.")
                break

    finally:
        remove_git_credentials()
