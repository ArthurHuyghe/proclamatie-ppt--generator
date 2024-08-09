import os
import shutil
import json

# Path to the JSON file
json_file_path = r"klaslijsten\klaslijsten zesde jaar.json"

# Path to the source directory where all student pictures are stored
source_directory = r"pasfotos\6e jaar 23-24\te sorteren"

# Path to the target directory where class folders should be created
target_directory = r"pasfotos\6e jaar 23-24"

# Load JSON data
with open(json_file_path, "r", encoding="utf-8") as file:
    data = json.load(file)


# Function to sanitize student names to match picture filenames
# def sanitize_name(name):
#    return name.lower().replace(" ", "_").replace("-", "_")


# Process each class group
for klasgroep in data["klasgroepen"]:
    klasgroep_name = klasgroep["klasgroep"]
    klasgroep_path = os.path.join(target_directory, klasgroep_name)

    # Create class folder if it doesn't exist
    if not os.path.exists(klasgroep_path):
        os.makedirs(klasgroep_path)

    # Move each student's picture to the corresponding class folder
    for leerling in klasgroep["leerlingen"]:
        student_name = leerling["Naam"]
        # sanitized_name = sanitize_name(student_name)
        picture_filename = f"{student_name}.jpg"
        source_picture_path = os.path.join(source_directory, picture_filename)
        target_picture_path = os.path.join(klasgroep_path, picture_filename)

        if os.path.exists(source_picture_path):
            shutil.move(source_picture_path, target_picture_path)
            print(f"Moved {picture_filename} to {klasgroep_path}")
        else:
            print(f"Picture not found for {student_name}: {picture_filename}")

print("Pictures have been organized into class folders.")
