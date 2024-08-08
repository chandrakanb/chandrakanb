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
    print("\nSetting up Git credentials...\n")
    # Ask for username and password
    username = input("Enter your Git username: ")
    password = getpass.getpass("Enter your Git password: ")

    try:
        # Configure Git to use credential helper store
        os.system("git config --global credential.helper store")
        
        # Get the full path to the home directory of the current user
        home_dir = os.path.expanduser("~")
        
        # Always create a new credentials file
        credentials_file = os.path.join(home_dir, ".git-credentials")
        with open(credentials_file, "w") as f:
            f.write(f"https://{username}:{password}@gitlab.kpit.com/ketakic/DRT_Jenkins.git\n")
        
        print("\nGit credentials configured and saved successfully.\n")
        return True
    except Exception as e:
        print(f"\nAn error occurred while setting up Git credentials: {e}\n")
        return False

if __name__ == "__main__":
    try:
        display_quote()
        
        # Set Git credentials
        set_git_credentials()
    except Exception as e:
        print(f"An error occurred: {e}")
