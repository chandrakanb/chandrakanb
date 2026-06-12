import sys


if len(sys.argv) == 2: 
    if len(sys.argv[1]) == 17 or len(sys.argv[1]) == 23:
        argument1 = sys.argv[1]
        commands = [command.upper() for command in [argument1]]
        print(commands)
        print(argument1)
    else:
        sys.exit("Error: hard_variant_code (17 chars) or fcp_switch_command (23 chars) are required.")
elif len(sys.argv) == 3:
    if len(sys.argv[1]) == 17 and len(sys.argv[2]) == 23:
        argument1 = sys.argv[1]
        argument2 = sys.argv[2]
        commands = [command.upper() for command in [argument1, argument2]]
        print(commands)
        print(commands[0])
        print(commands[1])
    elif len(sys.argv[1]) == 23 and len(sys.argv[2]) == 17:
        argument1 = sys.argv[2]
        argument2 = sys.argv[1]
        commands = [command.upper() for command in [argument1, argument2]]
        print(commands)
        print(commands[0])
        print(commands[1])
    else:
        sys.exit("Error: hard_variant_code (17 chars) and fcp_switch_command (23 chars) are required.")
    print("Argument1:", argument1)
    print("Argument2:", argument2)
else:
    sys.exit("Usage: python script.py <hard_variant_code>/<fcp_switch_command> or python script.py <hard_variant_code> <fcp_switch_command>")

#argument_commands = [command.upper() for command in commands]
#print(argument_commands)
"""
1
    smal
    large
2
saml and large

"""