for i in range (1, 2):
    # Let's run the Python script that was uploaded
    file_path = "D:\\py.py"

    # Executing the file
    with open(file_path) as file:
        script_content = file.read()

    exec(script_content)