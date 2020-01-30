def remove_none(list, second_list):
    for value in list:
        if value is None:
            list.remove(value)
        else:
            second_list.append(value)
