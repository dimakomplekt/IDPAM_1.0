import os

OLD_PART = "PROJECT_X"
NEW_PART = "PROJECT_Y"

# Корневая папка = папка, где лежит этот скрипт
root_dir = os.path.dirname(os.path.abspath(__file__))

for current_dir, _, files in os.walk(root_dir):
    for filename in files:

        if OLD_PART in filename:
            old_path = os.path.join(current_dir, filename)

            new_filename = filename.replace(OLD_PART, NEW_PART)
            new_path = os.path.join(current_dir, new_filename)

            try:
                os.rename(old_path, new_path)
                print(f"[OK] {filename} -> {new_filename}")
            except Exception as e:
                print(f"[ERROR] {filename}: {e}")

print("\nГотово.")
