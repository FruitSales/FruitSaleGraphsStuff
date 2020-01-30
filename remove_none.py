def remove_none(list, second_list):
    for value in list:
        if value is not None:
            second_list.append(value)
