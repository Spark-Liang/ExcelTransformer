import json

from xl_transform.common import Template
from xl_transform.reader import FileReader
from xl_transform.writer import FileWriter


def transfer(
        data_source_path,
        data_source_template_path,
        output_template_path,
        output_path,
        config_path
):
    with open(config_path, 'r') as f:
        config = json.loads("".join(f.readlines()))
    reader = FileReader(Template(data_source_template_path), config)
    writer = FileWriter(Template(output_template_path), config)
    data = reader.read(data_source_path)
    writer.write(data, output_path)
