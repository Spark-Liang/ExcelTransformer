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
