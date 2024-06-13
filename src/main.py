import argparse
import CPBE_first_page
import CPBC_all


def setup_arg_parser():
    """
    Set up the argument parser.
    """
    parser = argparse.ArgumentParser(description='parse for type of report')
    parser.add_argument('-product', '--product_report', action='store_true', help='If entered, return the algo report')
    parser.add_argument('-devops', '--devops_report', action='store_true', help='If entered, return the dev report')
    parser.add_argument('-algo', '--algo_report', action='store_true', help='If entered, return the algo report')
    parser.add_argument('-dev', '--dev_report', action='store_true', help='If entered, return the dev report')
    parser.add_argument('-bi', '--bi_report', action='store_true', help='If entered, return the bi report')
    parser.add_argument('-e', '--cpbe_report', action='store_true', help='If entered, return the  cpbe report')
    parser.add_argument('-input_file', '--input_file', type=str, help='Input file')
    parser.add_argument('-wd', '--wd', type=str, help='Working directory')
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
        if args.cpbe_report:
            print("Generating the CPBE report...")
            # Output file name for the full yearly report
            CPBE_first_page.main(args.wd, args.input_file)

        elif args.bi_report:
            print("Generating the Bi report...")
            CPBC_all.main(args.wd, args.input_file, "Bi")
        elif args.dev_report:
            print("Generating the Dev report...")
            CPBC_all.main(args.wd, args.input_file, "Dev")
        elif args.algo_report:
            print("Generating the Algo report...")
            CPBC_all.main(args.wd, args.input_file, "Algo")
        elif args.product_report:
            print("Generating the Product report...")
            CPBC_all.main(args.wd, args.input_file, "Product")
        elif args.devops_report:
            print("Generating the Devops report...")
            CPBC_all.main(args.wd, args.input_file, "Devops")
        else:
            print("please specify the type of report, use -h to know which types there are")

    except UnicodeDecodeError as e:
        print("An exception occurred:", str(e))


if __name__ == "__main__":
    main()
