import argparse

parser = argparse.ArgumentParser()
parser.add_argument('-p', '--ppt', '--pptx',
                    action='store_true', help='verbose flag')

args = parser.parse_args()

if args.ppt:
    print("~ Verbose!")
else:
    print("~ Not so verbose")
