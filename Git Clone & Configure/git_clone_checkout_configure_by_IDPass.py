import subprocess
import os
import sys
from urllib.parse import quote

def clone_repository(repo_url, clone_dir, username, password):
    """Clones the Git repository into the specified directory without specifying the branch."""
    try:
        print(f"Cloning repository {repo_url} into {clone_dir}.")
        
        # URL-encode the username and password
        encoded_username = quote(username)
        encoded_password = quote(password)
        
        # Set up the remote URL with the encoded credentials
        url_with_credentials = repo_url.replace("https://", f"https://{encoded_username}:{encoded_password}@")
        
        subprocess.run(["git", "clone", url_with_credentials, clone_dir], check=True)
        print(f"Repository cloned into {clone_dir}.")
    except subprocess.CalledProcessError as e:
        print(f"Error while cloning repository: {e}")
        raise

def checkout_branch(clone_dir, branch_name):
    """Checks out the specified branch in the cloned repository."""
    try:
        print(f"Checking out branch '{branch_name}' in {clone_dir}.")
        
        # Move to the cloned repository's directory
        os.chdir(clone_dir)
        
        # Checkout the specified branch
        subprocess.run(["git", "checkout", branch_name], check=True)
        print(f"Branch '{branch_name}' checked out successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Error while checking out branch: {e}")
        raise

def configure_git_token(clone_dir, username, password):
    """Configures Git with a username and password for authentication."""
    try:
        print(f"Configuring Git for authentication using the provided credentials in {clone_dir}.")
        
        # Move to the cloned repository's directory
        os.chdir(clone_dir)
        
        # Get the current remote URL
        remote_url_result = subprocess.run(["git", "remote", "get-url", "origin"], capture_output=True, text=True)
        remote_url = remote_url_result.stdout.strip()
        
        # URL-encode the username and password
        encoded_username = quote(username)
        encoded_password = quote(password)
        
        # Modify the remote URL to include the encoded username and password
        if remote_url.startswith("https://"):
            url_with_credentials = remote_url.replace("https://", f"https://{encoded_username}:{encoded_password}@")
            subprocess.run(["git", "remote", "set-url", "origin", url_with_credentials], check=True)
            print("Credentials configured successfully!")
        else:
            print("The repository doesn't use HTTPS, or the remote URL is invalid.")
    except subprocess.CalledProcessError as e:
        print(f"Error while configuring Git credentials: {e}")
        raise

if __name__ == "__main__":
    # Ensure the correct number of arguments are passed
    if len(sys.argv) != 6:
        print("Usage: python script_name.py <clone_directory> <repo_url> <branch> <username> <password>")
        sys.exit(1)

    # Get inputs from command-line arguments
    clone_directory = sys.argv[1]
    repository_url = sys.argv[2]
    branch_name = sys.argv[3]
    username = sys.argv[4]
    password = sys.argv[5]
    
    # Clone the repository
    clone_repository(repository_url, clone_directory, username, password)
    
    # Checkout the specified branch
    checkout_branch(clone_directory, branch_name)
    
    # Configure Git credentials in the specified directory
    configure_git_token(clone_directory, username, password)
