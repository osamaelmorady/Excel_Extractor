# scripts/init_generator.py
import os

def create_init_files(root="src"):
    for folder, dirs, files in os.walk(root):
        init_path = os.path.join(folder, '__init__.py')
        if '__init__.py' not in files:
            with open(init_path, 'w') as f:
                f.write("# Auto-generated __init__.py\n")
            print(f"✔ Created: {init_path}")

if __name__ == "__main__":
    create_init_files()