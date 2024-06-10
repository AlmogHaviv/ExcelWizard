import argparse
import CPBE_first_page
import CPBC_all


def setup_arg_parser():
    """
    Set up the argument parser.
    """
    parser = argparse.ArgumentParser(description='parse for type of report')
    parser.add_argument('-c', '--cpbc_report', action='store_true', help='If entered, return the cpbc report')
    parser.add_argument('-e', '--cpbe_report', action='store_true', help='If entered, return the  cpbe report')
    return parser

def main():
    try:
        """
        Main logic of the script using parsed arguments.
        """
        parser = setup_arg_parser()

        # Parse the arguments
        args = parser.parse_args()

        # Check which report type was selected and perform corresponding action
        if args.cpbc_report:
            print("Generating the CPBC report...")
            # Output file name for the full yearly report
            CPBE_first_page.main()

        elif args.cpbe_report:
            print("Generating the CPBE report...")
            CPBC_all.main()


        else:
            print("please specify the type of report, use -h to know which types there are")



    except UnicodeDecodeError as e:
        print("An exception occurred:", str(e))

