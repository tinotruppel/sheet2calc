var S2C = require('../'),
    _ = require('underscore'),
    files = ['test/test.xls',
        '../test_files/A4X_2013.xls',
        '../test_files/A4X_gnumeric.xls',
        '../test_files/AutoFilter.ods',
        '../test_files/AutoFilter.xls',
        '../test_files/AutoFilter.xlsb',
        '../test_files/AutoFilter.xlsx',
        '../test_files/BlankSheetTypes.ods',
        '../test_files/BlankSheetTypes.xls',
        '../test_files/BlankSheetTypes.xlsb',
        '../test_files/BlankSheetTypes.xlsm',
        //'../test_files/calendar_stress_test.xls',
        //'../test_files/calendar_stress_test.xlsb',
        //'../test_files/calendar_stress_test.xlsx',
        '../test_files/cell_style_simple.ods',
        '../test_files/cell_style_simple.xls',
        '../test_files/cell_style_simple.xlsb',
        '../test_files/cell_style_simple.xlsx',
        '../test_files/comments_stress_test.xls',
        '../test_files/comments_stress_test.xlsb',
        '../test_files/comments_stress_test.xlsx',
        '../test_files/custom_properties.xls',
        '../test_files/custom_properties.xlsb',
        '../test_files/custom_properties.xlsx',
        '../test_files/defined_names_simple.xls',
        '../test_files/defined_names_simple.xlsb',
        '../test_files/defined_names_simple.xlsx',
        '../test_files/ErrorTypes.xls',
        '../test_files/ErrorTypes.xlsb',
        '../test_files/ErrorTypes.xlsx',
        '../test_files/formulae_test_simple.xls',
        '../test_files/formulae_test_simple.xlsb',
        '../test_files/formulae_test_simple.xlsx',
        '../test_files/formula_stress_test.ods',
        '../test_files/formula_stress_test.xls',
        '../test_files/formula_stress_test.xlsb',
        '../test_files/formula_stress_test.xlsx',
        '../test_files/hyperlink_stress_test_2011.xls',
        '../test_files/hyperlink_stress_test_2011.xlsb',
        '../test_files/hyperlink_stress_test_2011.xlsx',
        '../test_files/interview.xlsx',
        '../test_files/issue.xlsx',
        '../test_files/large_strings.xls',
        '../test_files/large_strings.xlsb',
        '../test_files/large_strings.xlsx',
        '../test_files/LONumbers.xls',
        '../test_files/LONumbers.xlsx',
        '../test_files/LONumbers-2010.xls',
        '../test_files/LONumbers-2010.xlsx',
        '../test_files/LONumbers-2011.xls',
        '../test_files/LONumbers-2011.xlsx',
        '../test_files/merge_cells.ods',
        '../test_files/merge_cells.xls',
        '../test_files/merge_cells.xlsb',
        '../test_files/merge_cells.xlsx',
        '../test_files/mixed_sheets.xlsx',
        '../test_files/named_ranges_2011.xls',
        '../test_files/named_ranges_2011.xlsb',
        '../test_files/named_ranges_2011.xlsx',
        '../test_files/number_format.ods',
        '../test_files/number_format.xls',
        '../test_files/number_format.xlsb',
        '../test_files/number_format.xlsm',
        '../test_files/NumberFormatCondition.xls',
        '../test_files/NumberFormatCondition.xlsb',
        '../test_files/NumberFormatCondition.xlsm',
        '../test_files/number_format_entities.xls',
        '../test_files/number_format_entities.xlsb',
        '../test_files/number_format_entities.xlsx',
        '../test_files/number_format_russian.xls',
        '../test_files/number_format_russian.xlsb',
        '../test_files/number_format_russian.xlsm',
        '../test_files/numfmt_1_russian.xls',
        '../test_files/numfmt_1_russian.xlsb',
        '../test_files/numfmt_1_russian.xlsm',
        '../test_files/password_2002_40_972000.xls',
        '../test_files/password_2002_40_basecrypto.xls',
        '../test_files/password_2002_40_dhsc.xls',
        '../test_files/password_2002_40_dss.xls',
        '../test_files/password_2002_40_xor.xls',
        '../test_files/password_2002_128_enhanced.xls',
        '../test_files/password_2002_128_enhdss.xls',
        '../test_files/password_2002_128_enhrsa.xls',
        '../test_files/password_2002_128_enhrsasc.xls',
        '../test_files/password_2002_128_strong.xls',
        '../test_files/phonetic_text.xls',
        '../test_files/phonetic_text.xlsb',
        '../test_files/phonetic_text.xlsx',
        '../test_files/phpexcel_bad_cfb_dir.xls',
        '../test_files/pivot_table_named_range.xls',
        '../test_files/pivot_table_named_range.xlsb',
        '../test_files/pivot_table_named_range.xlsx',
        '../test_files/pivot_table_test.xls',
        '../test_files/pivot_table_test.xlsb',
        '../test_files/pivot_table_test.xlsm',
        '../test_files/rich_text_stress.ods',
        '../test_files/rich_text_stress.xls',
        '../test_files/rich_text_stress.xlsb',
        '../test_files/rich_text_stress.xlsx',
        '../test_files/RkNumber.xls',
        '../test_files/RkNumber.xlsb',
        '../test_files/RkNumber.xlsx',
        '../test_files/smart_tags_2007.xls',
        '../test_files/smart_tags_2007.xlsb',
        '../test_files/smart_tags_2007.xlsx',
        '../test_files/sushi.ods',
        '../test_files/sushi.xls',
        '../test_files/sushi.xlsb',
        '../test_files/sushi.xlsx',
        '../test_files/text_and_numbers.xls',
        '../test_files/text_and_numbers.xlsb',
        '../test_files/text_and_numbers.xlsx',
        '../test_files/time_stress_test_1.xlsb',
        '../test_files/write.xls',
        '../test_files/write.xlsx',
        '../test_files/xlsx-stream-d-date-cell.xls',
        '../test_files/xlsx-stream-d-date-cell.xlsb',
        '../test_files/xlsx-stream-d-date-cell.xlsx',
        '../test_files/חישוב_נקודות_זיכוי.xlsx'];

describe('Load and test sheets', function () {
    _.forEach(files, function (file) {
            it("Testing sheet '" + file + "'", function () {
                S2C.loadSheet(file);
            });
        }
    );
});

