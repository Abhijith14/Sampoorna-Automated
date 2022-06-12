def save_error(file, data):
    fo = open(file, "a")
    fo.write(str(data))
    fo.close()

# save_error("error.txt", ['f', 'e', 'd'])
# save_error("issue.txt", ['f', 'e', 'd'])
