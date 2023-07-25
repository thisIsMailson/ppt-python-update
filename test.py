
table_rows = {
    1: [9, 10], # if table has only one tuple
    2: [9, 10, 11, 12], # if table has two tuples
    3: [9, 10, 11, 12, 13, 14], # idem
    4: [9, 10, 11, 12, 13, 14, 45, 55],
    5: [9, 10, 11, 12, 13, 14, 45, 55, 58, 59]
}
table = [('concerned adv', 'this is the description'), ('les affair', 'the description for the lais affair')]
cells_to_populate = table_rows.get(len(table), [])


for i, cell in enumerate(cells_to_populate):
    cell_index = i // 2  # Integer division to get the corresponding tuple index

    if cell_index < len(table):
        if i % 2 == 0:  # Using i to determine if it's the first or second element of the tuple
            print(table[cell_index][0])  # First element of the tuple for column 1
        else:
            print(table[cell_index][1])  # Second element of the tuple for column 2


