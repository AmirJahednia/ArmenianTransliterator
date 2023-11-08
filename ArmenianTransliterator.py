import re
import pandas as pd


# A custom transliteration function for Armenian characters
def custom_armenian_transliterator(text):

    '''Dictionaries with mappings for Armenian characters, both upper and lower case and some additional rules 
    to take into account the rules of language, such as those depending on the position of the letter.''' 
    armenian_to_latin_lower = {
        u"ա": u"a",
        u"բ": u"b",
        u"գ": u"g",
        u"դ": u"d",
        u"ե": u"e",
        u"զ": u"z",
        u"է": u"e",
        u"ը": u"e",
        u"թ": u"t",
        u"ժ": u"zh",
        u"ի": u"i",
        u"լ": u"l",
        u"խ": u"kh",
        u"ծ": u"ts",
        u"կ": u"k",
        u"հ": u"h",
        u"ձ": u"dz",
        u"ղ": u"gh",
        u"ճ": u"ch",
        u"մ": u"m",
        u"յ": u"y",
        u"ն": u"n",
        u"շ": u"sh",
        u"ո": u"o",
        u"չ": u"ch",
        u"պ": u"p",
        u"ջ": u"j",
        u"ռ": u"r",
        u"ս": u"s",
        u"վ": u"v",
        u"տ": u"t",
        u"ր": u"r",
        u"ց": u"ts",
        u"ու": u"u",
        u"փ": u"p",
        u"ք": u"k",
        u"օ": u"o",
        u"ֆ": u"f",
    }

    armenian_to_latin_upper = {
        u"Ա": u"A",
        u"Բ": u"B",
        u"Գ": u"G",
        u"Դ": u"D",
        u"Ե": u"E",
        u"Զ": u"Z",
        u"Է": u"E",
        u"Ը": u"E",
        u"Թ": u"T",
        u"Ժ": u"Zh",
        u"Ի": u"I",
        u"Լ": u"L",
        u"Խ": u"Kh",
        u"Ծ": u"Ts",
        u"Կ": u"K",
        u"Ք": u"K",
        u"Հ": u"H",
        u"Ձ": u"Dz",
        u"Ղ": u"Gh",
        u"Ճ": u"Ch",
        u"Մ": u"M",
        u"Յ": u"Y",
        u"Ն": u"N",
        u"Շ": u"Sh",
        u"Ո": u"O",
        u"Չ": u"Ch",
        u"Պ": u"P",
        u"Փ": u"P",
        u"Ջ": u"J",
        u"Ռ": u"R",
        u"Ս": u"S",
        u"Վ": u"V",
        u"Տ": u"T",
        u"Ր": u"R",
        u"Ց": u"Ts",
        u"Ֆ": u"F",
        


    }

  
    # Check if the word starts with "Ե" or "Ե" and transliterate accordingly
    text = re.sub(u"^Ե", "Ye", text)  # If "Ե" is the first character, replace it with "Ye"
    text = re.sub(u"^ե", "ye", text)  # If "ե" is the first character, replace it with "ye"
    text = re.sub(u"(?<= )Ե", " Ye", text)  # Replace "ո" after a space with " Ye"
    text = re.sub(u"(?<= )ե", " ye", text)  # Replace "ո" after a space with " ye"    

    # Check if the word starts with "և"
    

    text = re.sub(r"^և", "Yev", text)  # Replace "և" at the start of a word with "Yev"
    text = re.sub(r" և ", " yev ", text)  # Replace "և" surrounded by spaces with "yev"
    text = re.sub(r"և", "ev", text)  # Replace all other occurrences of "և" with "ev"
    text = re.sub(r"եւ", "ev", text)  # Replace "եւ" with "ev"



    # Use a regular expression to find "ու" or "Ու" and transliterate accordingly
    text = re.sub(u"ու", "u", text)
    text = re.sub(u"Ու", "U", text)
    text = re.sub(u"ՈՒ", "U", text)


    # Check if the word starts with "Ո" or "ո" and transliterate accordingly
    text = re.sub(u"^Ո", "Vo", text)  # If "ո" is the first character, replace it with "Vo"
    text = re.sub(u"(?<= )ո", " vo", text)  # Replace "ո" after a space with " vo"

    # Transliterate the rest of the text using custom mapping
    transliterated_text = ""
    for char in text:
        if char in armenian_to_latin_lower:
            transliterated_text += armenian_to_latin_lower[char]
        elif char in armenian_to_latin_upper:
            transliterated_text += armenian_to_latin_upper[char]
        else:
            transliterated_text += char

    return transliterated_text



# Define a function to transliterate names with the option to exclude the last character in father names.
def transliterate_name(name, exclude_last_char=False):
    # Split the name into words for detection
    words = name.split()

    # Check if there are more than two words in the name
    if len(words) > 2:
        # If there are more than two words, process the last word
        if exclude_last_char and words[-1].endswith('յի'):
            name = name[:-2]  # Remove the last two characters 'yi'.
        elif exclude_last_char and words[-1].endswith('ի'):
            name = name[:-1]  # Remove the last character 'i'.
        elif exclude_last_char and words[-1].endswith('ու'):
            name = name[:-2] + 'ի'  # Remove ու and add 'ի'

    return custom_armenian_transliterator(name)

file_path = "Text.xlsx"  # path to your Excel file, if you change it please make the corresponding changes in both script and directory.

# The code that might face unexpected problems.
try:
    df = pd.read_excel(file_path, engine='openpyxl')  # Use 'openpyxl' as the engine

    # Transliterate the "Armenian" column to English using the custom function
    df['English'] = df['Armenian'].apply(lambda x: transliterate_name(str(x), exclude_last_char=False))

    # Save the modified Excel file
    df.to_excel(file_path, index=False, engine='openpyxl')

    print(f"Transliterated {len(df)} cells.")  # count all rows minus the first row which is the title.
    input('Press Enter to exit')

# Some minimal actions to prevent the program from crashing in case of an exception.
except FileNotFoundError:
    print(f"The file '{file_path}' was not found.")
    input('Press Enter to exit')

except pd.errors.ParserError:
    print("Error while parsing the Excel file.")
    input('Press Enter to exit')

except Exception as e:
    print(f"An error occurred: {e}")
    input('Press Enter to exit')