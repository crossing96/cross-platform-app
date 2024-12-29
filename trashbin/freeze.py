from py2exe import freeze

# Define the script details
console = [{
    'script': 'C:\\Users\\PC\\Desktop\\python\\pubxel\\pubxel.py' # path of the Python module of the executable target

}]

# Define the data files
data_files = [
    ('.', 'C:\\Users\\PC\\Desktop\\python\\pubxel\\impactfactor2022.txt'),
    ('.', 'C:\\Users\\PC\\Desktop\\python\\pubxel\\logo64.ico')
]

# Call the freeze function
freeze(console=console)
# freeze(console=console, data_files=data_files)