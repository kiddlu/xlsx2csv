#include <getopt.h>
#include <stdlib.h>
#include <string.h>
#include "xlsx2csv.h"

static void print_usage(const char *prog_name)
{
    printf("usage: %s [-h] [-v] [-a] [-c OUTPUTENCODING] [-d DELIMITER]\n", prog_name);
    printf("                [--hyperlinks] [-e] [--no-line-breaks]\n");
    printf("                [-E EXCLUDE_SHEET_PATTERN [EXCLUDE_SHEET_PATTERN ...]]\n");
    printf("                [-f DATEFORMAT] [-t TIMEFORMAT] [--floatformat FLOATFORMAT]\n");
    printf("                [--sci-float]\n");
    printf("                [-I INCLUDE_SHEET_PATTERN [INCLUDE_SHEET_PATTERN ...]]\n");
    printf("                [--exclude_hidden_sheets]\n");
    printf("                [--ignore-formats IGNORE_FORMATS [IGNORE_FORMATS ...]]\n");
    printf("                [-l LINETERMINATOR] [-m] [-n SHEETNAME] [-i]\n");
    printf("                [--skipemptycolumns] [-p SHEETDELIMITER] [-q QUOTING]\n");
    printf("                [-s SHEETID] [--include-hidden-rows]\n");
    printf("                xlsxfile [outfile]\n\n");
    printf("xlsx to csv converter\n\n");
    printf("positional arguments:\n");
    printf("  xlsxfile              xlsx file path, use '-' to read from STDIN\n");
    printf("  outfile               output csv file path\n\n");
    printf("options:\n");
    printf("  -h, --help            show this help message and exit\n");
    printf("  -v, --version         show program's version number and exit\n");
    printf("  -a, --all             export all sheets\n");
    printf("  -c, --outputencoding OUTPUTENCODING\n");
    printf("                        encoding of output csv (default: utf-8)\n");
    printf("  -d, --delimiter DELIMITER\n");
    printf("                        delimiter - columns delimiter in csv (default: ',')\n");
    printf("  --hyperlinks          include hyperlinks\n");
    printf("  -e, --escape          Escape \\r\\n\\t characters\n");
    printf("  --no-line-breaks      Replace \\r\\n\\t with space\n");
    printf("  -E, --exclude_sheet_pattern PATTERN\n");
    printf("                        exclude sheets matching pattern\n");
    printf("  -f, --dateformat DATEFORMAT\n");
    printf("                        override date/time format (ex. %%Y/%%m/%%d)\n");
    printf("  -t, --timeformat TIMEFORMAT\n");
    printf("                        override time format (ex. %%H/%%M/%%S)\n");
    printf("  --floatformat FLOATFORMAT\n");
    printf("                        override float format (ex. %%.15f)\n");
    printf("  --sci-float           force scientific notation to float\n");
    printf("  -I, --include_sheet_pattern PATTERN\n");
    printf("                        only include sheets matching pattern\n");
    printf("  --exclude_hidden_sheets\n");
    printf("                        Exclude hidden sheets from the output\n");
    printf("  -l, --lineterminator LINETERMINATOR\n");
    printf("                        line terminator (default: \\n)\n");
    printf("  -m, --merge-cells     merge cells\n");
    printf("  -n, --sheetname SHEETNAME\n");
    printf("                        sheet name to convert\n");
    printf("  -i, --ignoreempty     skip empty lines\n");
    printf("  --skipemptycolumns    skip trailing empty columns\n");
    printf("  -p, --sheetdelimiter SHEETDELIMITER\n");
    printf("                        sheet delimiter (default: '--------')\n");
    printf("  -q, --quoting QUOTING\n");
    printf("                        quoting mode: none, minimal, nonnumeric, all\n");
    printf("  -s, --sheet SHEETID   sheet number to convert\n");
    printf("  --include-hidden-rows include hidden rows\n");
}

int main(int argc, char **argv)
{
    xlsxOptions options     = {0};
    char       *infile      = NULL;
    char       *outfile     = NULL;
    int         sheetid     = 1;
    char       *sheetname   = NULL;
    bool        convert_all = false;

    /* Initialize default options */
    options.delimiter                   = ',';
    options.quoting                     = QUOTE_MINIMAL;
    options.sheetdelimiter              = "--------";
    options.dateformat                  = NULL;
    options.timeformat                  = NULL;
    options.floatformat                 = NULL;
    options.scifloat                    = false;
    options.skip_empty_lines            = false;
    options.skip_trailing_columns       = false;
    options.escape_strings              = false;
    options.no_line_breaks              = false;
    options.hyperlinks                  = false;
    options.include_sheet_pattern       = NULL;
    options.include_sheet_pattern_count = 0;
    options.exclude_sheet_pattern       = NULL;
    options.exclude_sheet_pattern_count = 0;
    options.exclude_hidden_sheets       = false;
    options.merge_cells                 = false;
    options.outputencoding              = "utf-8";
    options.lineterminator              = "\n";
    options.ignore_formats              = NULL;
    options.ignore_formats_count        = 0;
    options.skip_hidden_rows            = true;

    /* Parse command line options */
    static struct option long_options[] = {
        {"help",                  no_argument,       0, 'h' },
        {"version",               no_argument,       0, 'v' },
        {"all",                   no_argument,       0, 'a' },
        {"outputencoding",        required_argument, 0, 'c' },
        {"delimiter",             required_argument, 0, 'd' },
        {"hyperlinks",            no_argument,       0, 1001},
        {"escape",                no_argument,       0, 'e' },
        {"no-line-breaks",        no_argument,       0, 1002},
        {"exclude_sheet_pattern", required_argument, 0, 'E' },
        {"dateformat",            required_argument, 0, 'f' },
        {"timeformat",            required_argument, 0, 't' },
        {"floatformat",           required_argument, 0, 1003},
        {"sci-float",             no_argument,       0, 1004},
        {"include_sheet_pattern", required_argument, 0, 'I' },
        {"exclude_hidden_sheets", no_argument,       0, 1005},
        {"ignore-formats",        required_argument, 0, 1006},
        {"lineterminator",        required_argument, 0, 'l' },
        {"merge-cells",           no_argument,       0, 'm' },
        {"sheetname",             required_argument, 0, 'n' },
        {"ignoreempty",           no_argument,       0, 'i' },
        {"skipemptycolumns",      no_argument,       0, 1007},
        {"sheetdelimiter",        required_argument, 0, 'p' },
        {"quoting",               required_argument, 0, 'q' },
        {"sheet",                 required_argument, 0, 's' },
        {"include-hidden-rows",   no_argument,       0, 1008},
        {0,                       0,                 0, 0   }
    };

    int opt;
    while ((opt = getopt_long(argc, argv, "hvac:d:eE:f:t:I:l:mn:ip:q:s:", long_options, NULL)) !=
           -1) {
        switch (opt) {
            case 'h':
                print_usage(argv[0]);
                return 0;
            case 'v':
                printf("xlsx2csv %s\n", XLSX2CSV_VERSION);
                return 0;
            case 'a':
                convert_all = true;
                sheetid     = 0;
                break;
            case 'c':
                options.outputencoding = optarg;
                break;
            case 'd':
                if (strcmp(optarg, "tab") == 0 || strcmp(optarg, "\\t") == 0) {
                    options.delimiter = '\t';
                } else if (strlen(optarg) == 1) {
                    options.delimiter = optarg[0];
                } else if (optarg[0] == 'x' && strlen(optarg) > 1) {
                    options.delimiter = (char)strtol(optarg + 1, NULL, 16);
                } else {
                    fprintf(stderr, "Error: invalid delimiter\n");
                    return 1;
                }
                break;
            case 1001:
                options.hyperlinks = true;
                break;
            case 'e':
                options.escape_strings = true;
                break;
            case 1002:
                options.no_line_breaks = true;
                break;
            case 'E':
                /* TODO: Handle multiple patterns */
                break;
            case 'f':
                options.dateformat = optarg;
                break;
            case 't':
                options.timeformat = optarg;
                break;
            case 1003:
                options.floatformat = optarg;
                break;
            case 1004:
                options.scifloat = true;
                break;
            case 'I':
                /* TODO: Handle multiple patterns */
                break;
            case 1005:
                options.exclude_hidden_sheets = true;
                break;
            case 'l':
                if (strcmp(optarg, "\\n") == 0) {
                    options.lineterminator = "\n";
                } else if (strcmp(optarg, "\\r") == 0) {
                    options.lineterminator = "\r";
                } else if (strcmp(optarg, "\\r\\n") == 0) {
                    options.lineterminator = "\r\n";
                } else {
                    options.lineterminator = optarg;
                }
                break;
            case 'm':
                options.merge_cells = true;
                break;
            case 'n':
                sheetname = optarg;
                break;
            case 'i':
                options.skip_empty_lines = true;
                break;
            case 1007:
                options.skip_trailing_columns = true;
                break;
            case 'p':
                options.sheetdelimiter = optarg;
                break;
            case 'q':
                if (strcmp(optarg, "none") == 0) {
                    options.quoting = QUOTE_NONE;
                } else if (strcmp(optarg, "minimal") == 0) {
                    options.quoting = QUOTE_MINIMAL;
                } else if (strcmp(optarg, "nonnumeric") == 0) {
                    options.quoting = QUOTE_NONNUMERIC;
                } else if (strcmp(optarg, "all") == 0) {
                    options.quoting = QUOTE_ALL;
                } else {
                    fprintf(stderr, "Error: invalid quoting mode\n");
                    return 1;
                }
                break;
            case 's':
                sheetid = atoi(optarg);
                break;
            case 1008:
                options.skip_hidden_rows = false;
                break;
            default:
                print_usage(argv[0]);
                return 1;
        }
    }

    /* Get input and output files */
    if (optind >= argc) {
        fprintf(stderr, "Error: missing input file\n");
        print_usage(argv[0]);
        return 1;
    }

    infile = argv[optind++];
    if (optind < argc) {
        outfile = argv[optind];
    }

    /* Create converter */
    xlsx2csvConverter *conv = xlsx2csv_create(infile, &options);
    if (!conv) {
        fprintf(stderr, "Error: Failed to open %s\n", infile);
        return 1;
    }

    /* Convert */
    int result;
    if (convert_all) {
        if (!outfile) {
            fprintf(stderr, "Error: output directory required for --all\n");
            xlsx2csv_free(conv);
            return 1;
        }
        result = xlsx2csv_convert_all(conv, outfile);
    } else {
        result = xlsx2csv_convert(conv, outfile, sheetid, sheetname);
    }

    /* Cleanup */
    xlsx2csv_free(conv);

    return (result == 0) ? 0 : 1;
}
