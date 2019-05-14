import getopt
import sys

from xl_transform.reader import FileReader
from xl_transform.writer import FileWriter


def transfer(
        data_source_path,
        data_source_template_path,
        output_template_path,
        output_path,
        config_path=None
):
    extracted_mapping_data = FileReader.read(
        data_source_path,
        data_source_template_path,
        config_path
    )[0]
    FileWriter.write(
        output_path,
        output_template_path,
        extracted_mapping_data,
        config_path
    )


if __name__ == '__main__':
    usageStr = """xl_transform \\ 
                    --source or -s <source excel path>  \\
                    --source-template <source template file path> \\
                    --target-template <target template file path> \\
                    --target or -t <target excel file path> \\
                    --conf <control file path>
                """
    try:
        opts, args = getopt.getopt(sys.argv[1:], "s:t:",
                                   ["source=", "source-template=", "target-template=", "target=", "conf="])
    except getopt.GetoptError:
        print(usageStr)
        sys.exit(2)

    opts_dict = {}
    for opt, arg in opts:
        if opt[:2] == '--':
            opts_dict[opt[2:]] = arg
        elif opt[:1] == '-':
            opts_dict[opt[1]] = arg

    if "s" not in opts_dict and "source" not in opts_dict:
        raise Exception('please provide source file path via option "--source" or "-s"\n' + usageStr)
    if "t" not in opts_dict and "target" not in opts_dict:
        raise Exception('please provide target file path via option "--target" or "-t"\n' + usageStr)
    if "source-template" not in opts_dict:
        raise Exception('please provide source template file path via option "--source-template"\n' + usageStr)
    if "target-template" not in opts_dict:
        raise Exception('please provide target template file path via option "--target-template"\n' + usageStr)

    transfer(
        data_source_path=opts_dict["s"] if "s" in opts_dict else opts_dict["source"],
        data_source_template_path=opts_dict["source-template"],
        output_template_path=opts_dict["target-template"],
        output_path=opts_dict["t"] if "t" in opts_dict else opts_dict["target"],
        config_path=opts_dict["conf"] if "conf" in opts_dict else None
    )
