def launch_qfil(command):
    try:
        os.chdir("C:/Program Files (x86)/Qualcomm/QPST/bin/")
        with open("flashing_output.txt", "w") as output_file:
            result = subprocess.run(command, shell=True, stdout=output_file, stderr=subprocess.STDOUT, text=True)
        print("Flashing process completed. Output saved to 'flashing_output.txt'.")
    except subprocess.CalledProcessError as e:
        print(f"Error launching QFIL: {e}")
        sys.exit()
