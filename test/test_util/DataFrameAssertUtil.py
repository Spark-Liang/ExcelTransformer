from unittest import TestCase

import pandas as pd
from pandas import DataFrame


def assert_data_frame_with_same_header_and_data(
        test_case, expected_data_frame, result
):
    """

    :param test_case:
    :param DataFrame expected_data_frame:
    :param DataFrame result:
    :return:
    """
    test_case.assertEqual(
        expected_data_frame.shape, result.shape
    )
    for expected_header, result_header in zip(expected_data_frame.columns.values, result.columns.values):
        test_case.assertEqual(expected_header, result_header)
    max_x, max_y = expected_data_frame.shape
    for x in range(0, max_x):
        for y in range(0, max_y):
            test_case.assertEqual(
                expected_data_frame.iloc[x].iloc[y],
                result.iloc[x].iloc[y]
            )


def assert_data_in_excel_is_equal(
        test_case, expected_file_path, result_file_path, sheet_name=None
):
    """

    :param TestCase test_case:
    :param str expected_file_path:
    :param str result_file_path:
    :param str sheet_name: 
    :return:
    """
    expected_data = pd.read_excel(
        expected_file_path,
        sheet_name=sheet_name,
        header=None,
        dtype=str
    )
    result_data = pd.read_excel(
        result_file_path,
        sheet_name=sheet_name,
        header=None,
        dtype=str
    )

    # with single sheet in the excel
    if sheet_name is not None:
        assert_data_in_data_frame_is_equal(test_case, expected_data, result_data)
        return

    for sheet_name, expected_df in expected_data.items():
        test_case.assertTrue(sheet_name in result_data)
        result_df = result_data[sheet_name]
        assert_data_in_data_frame_is_equal(test_case, expected_df, result_df)


def assert_data_in_data_frame_is_equal(test_case, expected_df, result_df):
    test_case.assertEqual(
        expected_df.shape,
        result_df.shape
    )
    max_x, max_y = expected_df.shape
    for x in range(0, max_x):
        for y in range(0, max_y):
            expected_cell = expected_df.iloc[x].iloc[y]
            actual_cell = result_df.iloc[x].iloc[y]
            if pd.isna(expected_cell):
                continue
            test_case.assertEqual(
                expected_cell,
                actual_cell
            )
