import subprocess
import os
import argparse

def clone_repository(repo_url, branch_name, clone_dir):
    """Clones a specific branch of a Git repository into the specified directory."""
    try:
        print(f"Cloning branch '{branch_name}' from repository {repo_url} into {clone_dir}.")
        subprocess.run(["git", "clone", "--branch", branch_name, repo_url, clone_dir], check=True)
        print(f"Branch '{branch_name}' cloned into {clone_dir}.")
    except subprocess.CalledProcessError as e:
        print(f"Error while cloning repository: {e}")
        raise

def configure_git_token(clone_dir, token):
    """Configures Git with a personal access token for authentication."""
    try:
        print(f"Configuring Git for authentication using the provided token in {clone_dir}.")
        
        # Move to the cloned repository's directory
        os.chdir(clone_dir)
        
        # Get the current remote URL
        remote_url_result = subprocess.run(["git", "remote", "get-url", "origin"], capture_output=True, text=True)
        remote_url = remote_url_result.stdout.strip()
        
        # Modify the remote URL to include the token
        if remote_url.startswith("https://"):
            token_url = remote_url.replace("https://", f"https://{token}@")
            subprocess.run(["git", "remote", "set-url", "origin", token_url], check=True)
            print("Token configured successfully!")
        else:
            print("The repository doesn't use HTTPS, or the remote URL is invalid.")
    except subprocess.CalledProcessError as e:
        print(f"Error while configuring Git token: {e}")
        raise

if __name__ == "__main__":
    # User should provide the following values
    # Replace with the actual repository URL
    repository_url = "YOUR_REPOSITORY_URL"  
    # Replace with the desired branch name
    branch_name = "YOUR_BRANCH_NAME"  
    # Replace with your actual personal access token
    personal_access_token = "YOUR_PERSONAL_ACCESS_TOKEN"  

    # Parse the command-line argument for the clone directory
    parser = argparse.ArgumentParser(description="Specify the directory to clone the repository into.")
    parser.add_argument("clone_directory", help="The directory where the repository should be cloned.")

    args = parser.parse_args()

    # Clone the repository and configure the token in the specified directory
    clone_repository(repository_url, branch_name, args.clone_directory)
    configure_git_token(args.clone_directory, personal_access_token)
