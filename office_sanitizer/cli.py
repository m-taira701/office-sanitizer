import argparse
from excel import sanitize_excel

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("path")
    args = parser.parse_args()

    sanitize_excel(args.path)

if __name__ == "__main__":
    main()
