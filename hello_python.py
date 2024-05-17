def find(arr_list: list, func):
    for e in arr_list:
        if func(e):
            return e
    return None


if __name__ == '__main__':
    arr = ["1a", "2b", "3c", "4d", "5e"]
    # print(list(filter(lambda x: "" in x, arr))[0])
    # print(list(filter(lambda x: "c" in x, arr))[0])
    # print(find(arr, lambda x: "0" in x))
