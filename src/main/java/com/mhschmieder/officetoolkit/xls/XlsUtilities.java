/**
 * MIT License
 *
 * Copyright (c) 2020, 2025 Mark Schmieder
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 *
 * This file is part of the OfficeToolkit Library
 *
 * You should have received a copy of the MIT License along with the
 * OfficeToolkit Library. If not, see <https://opensource.org/licenses/MIT>.
 *
 * Project: https://github.com/mhschmieder/officetoolkit
 */
package com.mhschmieder.officetoolkit.xls;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheet;

import java.io.InputStream;
import java.util.List;

/**
 * This is a utility class for generating XLS spreadsheets and related objects,
 * using the Apache POI library.
 */
public final class XlsUtilities {

    /**
     * This method serves merely as a sanity check that the Maven integration
     * and builds work properly and also behave correctly inside Eclipse IDE. It
     * will likely get removed once I gain more confidence that I have solved
     * the well-known issues with Maven inside Eclipse as I move on to more
     * complex projects with dependencies (this project is quite simple and has
     * no dependencies at this time, until more functionality is added).
     *
     * @param args
     *            The command-line arguments for executing this class as the
     *            main entry point for an application
     *
     * @since 1.0
     */
    public static void main( final String[] args ) {
        System.out.println( "Hello Maven from OfficeToolkit!" ); //$NON-NLS-1$
    }

    /**
     * The default constructor is disabled, as this is a static utilities class.
     */
    private XlsUtilities() {}

    public static void addCategoryHeader( final Sheet sheet,
                                          final Row row,
                                          final int rowIndex,
                                          final int columnIndex,
                                          final int columnSpan,
                                          final String categoryName,
                                          final CellStyle categoryHeaderStyle ) {
        // Write the category header, which may span multiple columns. Use the
        // Apache POI Merge command for the spanning.
        final Cell categoryHeader = row.createCell( columnIndex, CellType.STRING );
        categoryHeader.setCellValue( categoryName );
        categoryHeader.setCellStyle( categoryHeaderStyle );
        sheet.addMergedRegion( new CellRangeAddress( rowIndex,
                                                     rowIndex,
                                                     columnIndex,
                                                     ( columnIndex + columnSpan ) - 1 ) );
    }

    public static void addCategoryHeaderRow( final Workbook workbook,
                                             final Sheet sheet,
                                             final int rowIndex,
                                             final int firstColumnIndex,
                                             final int[] columnSpans,
                                             final String[] categoryNames ) {
        // Get category header style for spanned column headers.
        final CellStyle categoryHeaderStyle = getColumnHeaderStyle( workbook );

        final Row row = sheet.createRow( rowIndex );

        // calculate the total number of columns, so we can work backwards.
        // TODO: Use new Java 8 functionality to sum the column spans.
        int numberOfColumns = 0;
        for ( final int columnSpan : columnSpans ) {
            numberOfColumns += columnSpan;
        }

        // Write the category headers, which may span multiple columns.
        // NOTE: It appears that addMergedRegion() might collapse the number of
        // columns vs. setting an actual column span value, so we process in
        // reverse order to see if all the category headers show up that way.
        int columnIndex = ( firstColumnIndex + numberOfColumns ) - 1;
        for ( int i = categoryNames.length - 1; i >= 0; i-- ) {
            final int columnSpan = columnSpans[ i ];
            final String categoryName = categoryNames[ i ];
            addCategoryHeader( sheet,
                               row,
                               rowIndex,
                               ( columnIndex - columnSpan ) + 1,
                               columnSpan,
                               categoryName,
                               categoryHeaderStyle );

            // Adjust the column index to the next category column.
            columnIndex -= columnSpan;
        }
    }

    public static void addCell( final Row row,
                                final int columnIndex,
                                final String cellValue,
                                final CellStyle cellStyle ) {
        final Cell columnHeader = row.createCell( columnIndex );
        columnHeader.setCellValue( cellValue );
        columnHeader.setCellStyle( cellStyle );
    }

    public static void addColumnHeaders( final Row row,
                                         final int firstColumnIndex,
                                         final String[] columnHeaderNames,
                                         final CellStyle columnHeaderStyle ) {
        // Write the column headers, which may only use one column each.
        int columnIndex = firstColumnIndex;
        for ( final String columnHeaderName : columnHeaderNames ) {
            addCell( row, columnIndex, columnHeaderName, columnHeaderStyle );
            columnIndex++;
        }
    }

    // This is a method to find the first column that matches an expected
    // column header, which is an essential part of parsing spreadsheets with
    // fully or partially known formatting and structure.
    public static int findFirstColumnToMatch( final Sheet sheet,
                                              final String columnHeaderToMatch,
                                              final int columnHeaderRowIndex,
                                              final int firstColumnToInspect ) {
        // Look at the header row, just to find the column pair for this trace.
        final Row headerRow = sheet.getRow( columnHeaderRowIndex );
        int firstColumnToMatch = firstColumnToInspect;
        final int numberOfColumns = headerRow.getPhysicalNumberOfCells();
        for ( int columnIndex =
                              firstColumnToInspect; columnIndex < numberOfColumns; columnIndex++ ) {
            final Cell cell = headerRow.getCell( columnIndex );
            final String columnHeader = cell.getStringCellValue();
            if ( columnHeaderToMatch.equals( columnHeader ) ) {
                firstColumnToMatch = columnIndex;
                break;
            }
        }

        return firstColumnToMatch;
    }

    // This is a convenience method and should not be used inside tight loops.
    public static Cell getColumnHeader( final Workbook workbook,
                                        final Row row,
                                        final int column,
                                        final String columnLabel ) {
        final Cell columnHeader = row.createCell( column );
        columnHeader.setCellValue( columnLabel );

        final CellStyle cellStyle = getColumnHeaderStyle( workbook );
        columnHeader.setCellStyle( cellStyle );

        return columnHeader;
    }

    public static CellStyle getColumnHeaderStyle( final Workbook workbook ) {
        // By default, style the cell with Black borders all around, and set an
        // Aqua background for the column header to stick out more.
        return getColumnHeaderStyle( workbook, IndexedColors.BLACK, IndexedColors.LIGHT_TURQUOISE );
    }

    public static CellStyle getColumnHeaderStyle( final Workbook workbook,
                                                  final IndexedColors foregroundColor,
                                                  final IndexedColors backgroundColor ) {
        final CellStyle cellStyle = workbook.createCellStyle();

        // Style the cell with thick borders all around.
        final BorderStyle borderStyle = BorderStyle.THICK;
        final short borderColorIndex = foregroundColor.getIndex();
        cellStyle.setBorderBottom( borderStyle );
        cellStyle.setBottomBorderColor( borderColorIndex );
        cellStyle.setBorderLeft( borderStyle );
        cellStyle.setLeftBorderColor( borderColorIndex );
        cellStyle.setBorderRight( borderStyle );
        cellStyle.setRightBorderColor( borderColorIndex );
        cellStyle.setBorderTop( borderStyle );
        cellStyle.setTopBorderColor( borderColorIndex );

        // Try to make sure the headers don't jam up right against the borders.
        cellStyle.setIndention( ( short ) 1 );

        // NOTE: Must set the foreground color to the intended background, and
        // then set a solid foreground fill pattern, to achieve colored cells.
        cellStyle.setFillForegroundColor( backgroundColor.getIndex() );
        cellStyle.setFillPattern( FillPatternType.SOLID_FOREGROUND );

        // Use a bold, italic font, larger than for the data cells.
        final Font font = workbook.createFont();
        font.setFontHeightInPoints( ( short ) 12 );
        font.setFontName( "Arial" ); //$NON-NLS-1$
        font.setItalic( true );
        font.setBold( true );
        cellStyle.setFont( font );

        // Column headers are often wider than data values, so it is best to
        // wrap text so that they are visible without much effort by the end
        // user (i.e. they shouldn't have to stretch column widths every time,
        // or even the first time, they open a document).
        cellStyle.setWrapText( true );

        // Align headers in the center of the cell, as this especially helps
        // when there is either a big difference between header length and
        // content length (in terms of number of characters) or when columns are
        // merged for column-span category headers.
        cellStyle.setAlignment( HorizontalAlignment.CENTER );

        return cellStyle;
    }

    // Get the Workbook from an Input Stream, accounting for Spreadsheet Format,
    // as well as the names of sheets that must be treated for large data vs.
    // using DOM (in order to avoid Out of Memory run-time exceptions).
    @SuppressWarnings("nls")
    public static Workbook getWorkbook( final InputStream inputStream,
                                        final String spreadsheetFormat,
                                        final boolean useBigDataStrategy,
                                        final List< String > largeSheetNames ) {
        Workbook workbook = null;

        try {
            switch ( spreadsheetFormat ) {
            case "xlsx":
                // NOTE: In this context, it is too hard to revert the document
                // once done, and the performance is no better than using the
                // direct approach via an Input Stream anyway.
                // final OPCPackage pkg = OPCPackage.open( inputStream );
                workbook = // new XSSFWorkbook( pkg )
                         new XSSFWorkbook( inputStream ) {
                             /**
                              * Avoid DOM parsing of large sheets.
                              */
                             @Override
                             public void parseSheet( final java.util.Map< String, XSSFSheet > shIdMap,
                                                     final CTSheet ctSheet ) {
                                 // Skip parsing this sheet is using the Big
                                 // Data strategy, or if the sheet is not
                                 // found in a pre-curated list of large
                                 // sheets, as DOM is expensive for Big data.
                                 if ( !useBigDataStrategy || ( largeSheetNames == null )
                                         || largeSheetNames.isEmpty()
                                         || !largeSheetNames.contains( ctSheet.getName() ) ) {
                                     super.parseSheet( shIdMap, ctSheet );
                                 }
                             }
                         };
                break;
            case "xls":
                workbook = new HSSFWorkbook( inputStream );
                break;
            default:
                break;
            }
        }
        catch ( final Exception e ) {
            e.printStackTrace();
        }

        return workbook;
    }

    // Get the Workbook for an Output Stream, accounting for Spreadsheet Format
    // as well as whether to use a Big Data strategy to conserve heap space.
    @SuppressWarnings("nls")
    public static Workbook getWorkbook( final String spreadsheetFormat,
                                        final boolean useBigDataStrategy ) {
        Workbook workbook = null;
        switch ( spreadsheetFormat ) {
        case "xlsx":
            workbook = useBigDataStrategy ? new SXSSFWorkbook() : new XSSFWorkbook();
            break;
        case "xls":
            workbook = new HSSFWorkbook();
            break;
        default:
            break;
        }

        return workbook;
    }

}
