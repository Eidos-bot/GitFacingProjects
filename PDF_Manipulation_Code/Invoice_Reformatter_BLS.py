# What's going on here is any X in the inv template will remove the character in the input string thatt matches it's location. sa
# If you put a period you are adding a period. The 0s are just used as placed holders. They are ignored.
# Example:
# input_string = "$S052856048.003"
# inv_template = "XX00000000OXXO"

def invoice_adj(raw_inv, inv_template):
    raw_inv_1A = raw_inv.replace("S$", "S").replace("$S", "S").replace("$", "S").replace("SS", "S").replace("/", "7").replace("£", "E").replace("¢", "C")
    print(raw_inv_1A)
    new_inv = list(raw_inv_1A)  # Convert the input string to a list
    for index, char in enumerate(inv_template):
        if index < len(new_inv):
            if char == "X":
                new_inv[index] = ''
            if char == ".":
                new_inv.insert(index, ".")
    print(f"{''.join(new_inv)}")
    return ''.join(new_inv)  # Convert the modified list back to a string

# invoice_adj(input_string, inv_template)

