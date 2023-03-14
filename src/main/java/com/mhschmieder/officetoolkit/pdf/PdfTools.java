/**
 * MIT License
 *
 * Copyright (c) 2020, 2022 Mark Schmieder
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
package com.mhschmieder.officetoolkit.pdf;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import javax.imageio.ImageIO;

import org.apache.commons.math3.util.FastMath;

import com.mhschmieder.commonstoolkit.branding.ProductBranding;
import com.mhschmieder.commonstoolkit.io.FileUtilities;
import com.pdfjet.Align;
import com.pdfjet.Cell;
import com.pdfjet.Color;
import com.pdfjet.Compliance;
import com.pdfjet.Font;
import com.pdfjet.Image;
import com.pdfjet.ImageType;
import com.pdfjet.Letter;
import com.pdfjet.PDF;
import com.pdfjet.Page;
import com.pdfjet.PageLayout;
import com.pdfjet.PageMode;
import com.pdfjet.Paragraph;
import com.pdfjet.Point;
import com.pdfjet.Table;
import com.pdfjet.TextFrame;
import com.pdfjet.TextLine;

/**
 * This class has a set of tools that are useful for minimizing cut/paste work
 * for common PDF tasks such as table creation, cell formatting, etc.
 * <p>
 * TODO: Split this class in two, for Project Report specific helper methods?
 * <p>
 * NOTE: This code was originally based off of a commercial license for the 5.9
 * release, which contains a TextColumn class for aligning and spacing multiple
 * Paragraphs on a Page. As this OfficeToolkit is now an open source library, it
 * must instead be based on the free 5.75 release of pdfJet, so we now use the
 * TextFrame class. THE OUTPUT HAS NOT YET BEEN TESTED, as my GUI app is not
 * currently in a fully buildable state.
 * <p>
 * TODO: Consider switching to Apache PdfBox in conjunction with an add-on
 * library such as easytable, ph-pdf-layout, pdfbox-layout, or PdfLayoutManager,
 * as pdfJet appears to have been abandoned at the start of the COVID pandemic
 * in early 2020 and as PdfBox and its add-on high-level layout libraries are
 * growing quickly in functionality, almost matching or surpassing iText.
 */
public final class PdfTools {

    // Declare the default point size per inch.
    public static final float   POINTS_PER_INCH              = 72f;

    // Use Letter page size as most common case, for scaling/clipping/etc.
    public static final float[] PORTRAIT_PAGE_SIZE           = Letter.PORTRAIT;
    public static final float[] LANDSCAPE_PAGE_SIZE          = Letter.LANDSCAPE;

    // Declare constants for page margins, using default page size assumption.
    public static final float   PORTRAIT_LEFT_MARGIN         = 0.75f * POINTS_PER_INCH;
    public static final float   PORTRAIT_TOP_MARGIN          = 0.75f * POINTS_PER_INCH;
    public static final float   PORTRAIT_RIGHT_MARGIN        = 0.5f * POINTS_PER_INCH;
    public static final float   PORTRAIT_BOTTOM_MARGIN       = 1f * POINTS_PER_INCH;

    public static final float   LANDSCAPE_LEFT_MARGIN        = 0.75f * POINTS_PER_INCH;
    public static final float   LANDSCAPE_TOP_MARGIN         = 0.75f * POINTS_PER_INCH;
    public static final float   LANDSCAPE_RIGHT_MARGIN       = 0.5f * POINTS_PER_INCH;
    public static final float   LANDSCAPE_BOTTOM_MARGIN      = 1f * POINTS_PER_INCH;

    // We only need to compute the resulting layout size once.
    public static final float   PORTRAIT_PAGE_LAYOUT_WIDTH   = PORTRAIT_PAGE_SIZE[ 0 ]
            - PORTRAIT_LEFT_MARGIN - PORTRAIT_RIGHT_MARGIN;
    public static final float   PORTRAIT_PAGE_LAYOUT_HEIGHT  = PORTRAIT_PAGE_SIZE[ 1 ]
            - PORTRAIT_TOP_MARGIN - PORTRAIT_BOTTOM_MARGIN;

    public static final float   LANDSCAPE_PAGE_LAYOUT_WIDTH  = LANDSCAPE_PAGE_SIZE[ 0 ]
            - LANDSCAPE_LEFT_MARGIN - LANDSCAPE_RIGHT_MARGIN;
    public static final float   LANDSCAPE_PAGE_LAYOUT_HEIGHT = LANDSCAPE_PAGE_SIZE[ 1 ]
            - LANDSCAPE_TOP_MARGIN - LANDSCAPE_BOTTOM_MARGIN;

    @SuppressWarnings("nls")
    public static void addEmptyLines( final List< Paragraph > paragraphs,
                                      final int alignment,
                                      final Font font,
                                      final int numberOfLines ) {
        for ( int i = 0; i < numberOfLines; i++ ) {
            addParagraph( paragraphs, alignment, font, "" );
        }
    }

    public static void addParagraph( final List< Paragraph > paragraphs,
                                     final int alignment,
                                     final Font font,
                                     final String text ) {
        final TextLine textLine = new TextLine( font, text );
        final Paragraph paragraph = new Paragraph();
        paragraph.setAlignment( alignment );
        paragraph.add( textLine );
        paragraphs.add( paragraph );
    }

    // Utility method to add a stylized table cell, with custom justification.
    public static void addTableCell( final List< Cell > rowData,
                                     final Font font,
                                     final int backgroundColor,
                                     final int textFillColor,
                                     final int align,
                                     final String cellValue ) {
        addTableCell( rowData, 1, font, backgroundColor, textFillColor, align, true, cellValue );
    }

    // Utility method to add a stylized table cell, with center justification.
    public static void addTableCell( final List< Cell > rowData,
                                     final Font font,
                                     final int backgroundColor,
                                     final int textFillColor,
                                     final String cellValue ) {
        addTableCell( rowData, 1, font, backgroundColor, textFillColor, cellValue );
    }

    // Utility method to add a stylized table cell, with custom justification.
    public static void addTableCell( final List< Cell > rowData,
                                     final int columnSpan,
                                     final Font font,
                                     final int backgroundColor,
                                     final int textFillColor,
                                     final int align,
                                     final boolean paintBorders,
                                     final String cellValue ) {
        // Note that pen color is for lines, and brush color is for text.
        final Cell tableCell = new Cell( font, cellValue );

        tableCell.setColSpan( columnSpan );
        tableCell.setBgColor( backgroundColor );
        tableCell.setPenColor( Color.black );
        tableCell.setBrushColor( textFillColor );
        tableCell.setTextAlignment( align );
        tableCell.setPadding( 4f );

        if ( !paintBorders ) {
            tableCell.setNoBorders();
            tableCell.setBottomPadding( 0f );
        }

        rowData.add( tableCell );
    }

    // Utility method to add a stylized table cell, with center justification.
    public static void addTableCell( final List< Cell > rowData,
                                     final int columnSpan,
                                     final Font font,
                                     final int backgroundColor,
                                     final int textFillColor,
                                     final String cellValue ) {
        addTableCell( rowData,
                      columnSpan,
                      font,
                      backgroundColor,
                      textFillColor,
                      Align.CENTER,
                      true,
                      cellValue );
    }

    // Utility method to add a stylized table cell, with custom justification.
    public static void addTableCell( final List< Cell > rowData,
                                     final PdfFonts fonts,
                                     final int align,
                                     final boolean paintBorders,
                                     final String cellValue ) {
        addTableCell( rowData,
                      1,
                      fonts._tableCellFont,
                      Color.white,
                      Color.black,
                      align,
                      paintBorders,
                      cellValue );
    }

    // Utility method to add a stylized table cell, with custom justification.
    public static void addTableCell( final List< Cell > rowData,
                                     final PdfFonts fonts,
                                     final int backgroundColor,
                                     final int textFillColor,
                                     final int align,
                                     final String cellValue ) {
        addTableCell( rowData,
                      fonts._tableCellFont,
                      backgroundColor,
                      textFillColor,
                      align,
                      cellValue );
    }

    // Utility method to add a stylized table cell, with center justification.
    public static void addTableCell( final List< Cell > rowData,
                                     final PdfFonts fonts,
                                     final int backgroundColor,
                                     final int textFillColor,
                                     final String cellValue ) {
        addTableCell( rowData, fonts._tableCellFont, backgroundColor, textFillColor, cellValue );
    }

    // Utility method to add a stylized table cell, with custom justification.
    public static void addTableCell( final List< Cell > rowData,
                                     final PdfFonts fonts,
                                     final int align,
                                     final String cellValue ) {
        addTableCell( rowData, fonts, align, true, cellValue );
    }

    // Utility method to add a stylized table cell, with center justification.
    public static void addTableCell( final List< Cell > rowData,
                                     final PdfFonts fonts,
                                     final String cellValue ) {
        addTableCell( rowData, fonts, Align.CENTER, cellValue );
    }

    // Utility method to avoid indexing errors or null cell pointers in sparse
    // tables. Not used currently, but should be applied to all tables.
    public static void appendMissingCells( final List< List< Cell > > tableData,
                                           final Font tableCellFont ) {
        final List< Cell > firstRow = tableData.get( 0 );
        final int numOfColumns = firstRow.size();
        for ( final List< Cell > dataRow : tableData ) {
            final int dataRowColumns = dataRow.size();
            if ( dataRowColumns < numOfColumns ) {
                for ( int j = 0; j < ( numOfColumns - dataRowColumns ); j++ ) {
                    dataRow.add( new Cell( tableCellFont ) );
                }
                dataRow.get( dataRowColumns - 1 )
                        .setColSpan( ( numOfColumns - dataRowColumns ) + 1 );
            }
        }
    }

    // Generic method to create PDF-ready Information Table Data.
    public static List< List< Cell > > createInformationTableData( final PdfFonts borderlessTableFonts,
                                                                   final int align,
                                                                   final String[] information )
            throws Exception {
        // Information tables only have one column, with multiple rows.
        final List< List< Cell > > informationTableData = new ArrayList<>();

        // Push all of the Information fields to unique rows (borderless).
        for ( final String element : information ) {
            final List< Cell > informationRowData = new ArrayList<>();
            PdfTools.addTableCell( informationRowData,
                                   borderlessTableFonts,
                                   align,
                                   false,
                                   element );
            informationTableData.add( informationRowData );
        }

        return informationTableData;
    }

    // Utility method to create a stylized table with the given column names.
    public static Table createTable( final List< List< Cell > > tableData,
                                     final PdfFonts fonts,
                                     final String[] columnNames,
                                     final boolean landscapeMode ) {
        return createTable( tableData, fonts, columnNames, null, landscapeMode );
    }

    // Utility method to create a stylized table with the given column names.
    public static Table createTable( final List< List< Cell > > tableData,
                                     final PdfFonts fonts,
                                     final String[] columnNames,
                                     final int[] columnWidthsInPixels,
                                     final boolean landscapeMode ) {
        return createTable( tableData,
                            fonts,
                            null,
                            null,
                            columnNames,
                            columnNames.length,
                            columnWidthsInPixels,
                            landscapeMode );
    }

    // Utility method to create a labeled table with the given column names.
    public static Table createTable( final List< List< Cell > > tableData,
                                     final PdfFonts fonts,
                                     final String[] spanNames,
                                     final int[] spanLengths,
                                     final int labelBackgroundColor,
                                     final int labelTextFillColor,
                                     final String[] columnNames,
                                     final int numberOfColumns,
                                     final boolean landscapeMode ) {
        return createTable( tableData,
                            fonts,
                            spanNames,
                            spanLengths,
                            labelBackgroundColor,
                            labelTextFillColor,
                            columnNames,
                            numberOfColumns,
                            null,
                            landscapeMode );
    }

    // Utility method to create a labeled table with the given column names.
    @SuppressWarnings("nls")
    public static Table createTable( final List< List< Cell > > tableData,
                                     final PdfFonts fonts,
                                     final String[] spanNames,
                                     final int[] spanLengths,
                                     final int labelBackgroundColor,
                                     final int labelTextFillColor,
                                     final String[] columnNames,
                                     final int numberOfColumns,
                                     final int[] columnWidthsInPixels,
                                     final boolean landscapeMode ) {
        final Table table = new Table();

        if ( ( spanNames == null ) && ( columnNames == null ) ) {
            return table;
        }

        // Determine the scale factor for column widths, which are expressed in
        // pixels but need to be recalculated in points based on the page layout
        // width (which accounts for margins, and landscape vs. portrait mode).
        final float columnWidthScaleFactor = getColumnWidthScaleFactor( columnWidthsInPixels,
                                                                        landscapeMode );

        // Write the span names for the table.
        // TODO: Even though it's unlikely we'll have manual column widths in
        // tables that also have a full-span label or two at the top, we should
        // write an algorithm or enhanced table cell methods that account for
        // scaled column width preferences when present (i.e. nun-null).
        if ( ( spanNames != null ) && ( spanLengths != null ) ) {
            final List< Cell > spanHeaders = new ArrayList<>();

            int spanLengthIndex = 0;
            int columnIndex = 0;
            int columnWidthIndex = 0;
            int columnBlankStartIndex = 1;
            int columnBlankEndIndex = 1;
            for ( final String spanName : spanNames ) {
                final int columnSpan = spanLengths[ spanLengthIndex ];
                addTableCell( spanHeaders,
                              columnSpan,
                              fonts._tableLabelFont,
                              labelBackgroundColor,
                              labelTextFillColor,
                              spanName );

                // If manual column widths are present, set the current column
                // width now, after scaling from pixels to points.
                // TODO: Modify the addTableCell method to do this?
                if ( columnWidthsInPixels != null ) {
                    float columnSpanWidth = 0f;
                    for ( int i = 0; i < columnSpan; i++ ) {
                        final float columnWidth = columnWidthsInPixels[ columnWidthIndex++ ]
                                * columnWidthScaleFactor;
                        columnSpanWidth += columnWidth;
                    }
                    spanHeaders.get( columnIndex ).setWidth( columnSpanWidth );
                }

                columnIndex++;

                spanLengthIndex++;

                // NOTE: Due to a bug (or weird design decision) in PDFjet, we
                // have to blank-fill the unused columns, or we get run-time
                // indexing crashes thanks to look-ahead to non-existent columns
                // in the PDFjet source code.
                columnBlankEndIndex = ( spanLengthIndex >= spanLengths.length )
                    ? numberOfColumns - 1
                    : ( columnBlankStartIndex + columnSpan ) - 1 - 1;
                for ( int columnBlankIndex =
                                           columnBlankStartIndex; columnBlankIndex <= columnBlankEndIndex; columnBlankIndex++ ) {
                    addTableCell( spanHeaders,
                                  fonts._tableHeaderFont,
                                  labelBackgroundColor,
                                  labelTextFillColor,
                                  "" );

                    // Make sure the blankers don't add any overall width.
                    spanHeaders.get( columnIndex ).setWidth( 0f );

                    columnIndex++;
                }

                columnBlankStartIndex += columnSpan;
            }

            tableData.add( spanHeaders );
        }

        // Write the column names for the table.
        // TODO: Even though it's unlikely we'll have manual column widths in
        // tables that have any multi-span column headers, we should scale the
        // column width preferences when present (i.e. nun-null).
        if ( columnNames != null ) {
            final List< Cell > columnHeaders = new ArrayList<>();

            if ( numberOfColumns == columnNames.length ) {
                int columnWidthIndex = 0;
                for ( final String columnName : columnNames ) {
                    addTableCell( columnHeaders,
                                  fonts._tableHeaderFont,
                                  PdfColors.TABLE_HEADER_BACKGROUND_COLOR,
                                  PdfColors.TABLE_HEADER_TEXT_FILL_COLOR,
                                  columnName );

                    // If manual column widths are present, set the current
                    // column width now, after scaling from pixels to points.
                    // TODO: Modify the addTableCell method to do this?
                    if ( columnWidthsInPixels != null ) {
                        final float columnWidth = columnWidthsInPixels[ columnWidthIndex ]
                                * columnWidthScaleFactor;
                        columnHeaders.get( columnWidthIndex ).setWidth( columnWidth );
                        columnWidthIndex++;
                    }
                }
            }
            else {
                // Deal with crude current algorithm that only handles the final
                // column header spanning multiple columns.
                final int numberOfSingleColumnHeaders = columnNames.length - 1;
                for ( int i = 0; i < numberOfSingleColumnHeaders; i++ ) {
                    addTableCell( columnHeaders,
                                  fonts._tableHeaderFont,
                                  PdfColors.TABLE_HEADER_BACKGROUND_COLOR,
                                  PdfColors.TABLE_HEADER_TEXT_FILL_COLOR,
                                  columnNames[ i ] );
                }

                final int multiColumnHeaderIndex = numberOfSingleColumnHeaders;
                addTableCell( columnHeaders,
                              numberOfColumns - numberOfSingleColumnHeaders,
                              fonts._tableHeaderFont,
                              PdfColors.TABLE_HEADER_BACKGROUND_COLOR,
                              PdfColors.TABLE_HEADER_TEXT_FILL_COLOR,
                              columnNames[ multiColumnHeaderIndex ] );

                // NOTE: Due to a bug (or weird design decision) in PDFjet, we
                // have to blank-fill the unused columns, or we get run-time
                // indexing crashes thanks to look-ahead to non-existent columns
                // in the PDFjet source code.
                for ( int i = multiColumnHeaderIndex + 1; i < numberOfColumns; i++ ) {
                    addTableCell( columnHeaders,
                                  fonts._tableHeaderFont,
                                  PdfColors.TABLE_HEADER_BACKGROUND_COLOR,
                                  PdfColors.TABLE_HEADER_TEXT_FILL_COLOR,
                                  "" );
                }
            }

            tableData.add( columnHeaders );
        }

        return table;
    }

    // Utility method to create a labeled table with the given column names.
    public static Table createTable( final List< List< Cell > > tableData,
                                     final PdfFonts fonts,
                                     final String[] spanNames,
                                     final int[] spanLengths,
                                     final String[] columnNames,
                                     final int numberOfColumns,
                                     final boolean landscapeMode ) {
        return createTable( tableData,
                            fonts,
                            spanNames,
                            spanLengths,
                            columnNames,
                            numberOfColumns,
                            null,
                            landscapeMode );
    }

    // Utility method to create a labeled table with the given column names.
    public static Table createTable( final List< List< Cell > > tableData,
                                     final PdfFonts fonts,
                                     final String[] spanNames,
                                     final int[] spanLengths,
                                     final String[] columnNames,
                                     final int numberOfColumns,
                                     final int[] columnWidthsInPixels,
                                     final boolean landscapeMode ) {
        return createTable( tableData,
                            fonts,
                            spanNames,
                            spanLengths,
                            PdfColors.TABLE_LABEL_BACKGROUND_COLOR,
                            PdfColors.TABLE_LABEL_TEXT_FILL_COLOR,
                            columnNames,
                            numberOfColumns,
                            columnWidthsInPixels,
                            landscapeMode );
    }

    public static Table createTable( final List< List< Cell > > tableData,
                                     final PdfFonts fonts,
                                     final String[] spanNames,
                                     final int[] spanLengths,
                                     final String[] columnNames,
                                     final int[] columnWidthsInPixels,
                                     final boolean landscapeMode ) {
        return createTable( tableData,
                            fonts,
                            spanNames,
                            spanLengths,
                            columnNames,
                            columnNames.length,
                            columnWidthsInPixels,
                            landscapeMode );
    }

    public static float getColumnWidthScaleFactor( final int[] columnWidthsInPixels,
                                                   final boolean landscapeMode ) {
        // First calculate the total table width, expressed in pixels.
        // TODO: Take advantage of new Java 8 functional programming features.
        float tableWidthPixels = 0f;
        if ( columnWidthsInPixels != null ) {
            for ( final int columnWidth : columnWidthsInPixels ) {
                tableWidthPixels += columnWidth;
            }
        }

        // Determine the scale factor for column widths, which are expressed in
        // pixels but need to be recalculated in points based on the page layout
        // width (which accounts for margins).
        final float pageWidth = landscapeMode
            ? LANDSCAPE_PAGE_LAYOUT_WIDTH
            : PORTRAIT_PAGE_LAYOUT_WIDTH;
        final float columnWidthScaleFactor = ( tableWidthPixels > 0f )
            ? pageWidth / tableWidthPixels
            : 1f;

        return columnWidthScaleFactor;
    }

    // This is the main method to get a PDF document from PDFjet.
    public static PDF getDocument( final OutputStream outputStream,
                                   final ProductBranding productBranding,
                                   final String reportSubtitle,
                                   final String projectName,
                                   final String designer )
            throws Exception {
        // Generate the report file title for use within the PDF report.
        final String reportTitle = getReportTitle( productBranding, reportSubtitle );

        // Set the main subject header, as the project name.
        final String reportSubject = projectName;

        // Note that PDFjet uses the author to set the creator field.
        final String reportAuthor = designer;

        final PDF document = getDocument( outputStream, reportTitle, reportSubject, reportAuthor );

        return document;
    }

    // This is the main method to get a PDF document from PDFjet.
    public static PDF getDocument( final OutputStream outputStream,
                                   final String reportTitle,
                                   final String reportSubject,
                                   final String reportAuthor )
            throws Exception {
        // NOTE: We specify PDF/A compliance, which also means that all fonts
        // must be embedded in the PDF document.
        final PDF document = new PDF( outputStream, Compliance.PDF_A_1B );

        // As there is no way to set multiple Page Modes, prioritize Outlines.
        document.setPageMode( PageMode.USE_OUTLINES );

        // For Reports, we tend to want single-column layout vs. newspaper-style
        // dual-column, as most of the report will be tables, images, and other
        // objects vs. a text document content bias.
        document.setPageLayout( PageLayout.ONE_COLUMN );

        // Write the file title for the PDF report.
        document.setTitle( reportTitle );

        // Write the main subject header, as the project name.
        document.setSubject( reportSubject );

        // Note that PDFjet uses the author to set the creator field.
        document.setAuthor( reportAuthor );

        return document;
    }

    @SuppressWarnings("nls")
    public static Image getImageSnapshot( final PDF document, final BufferedImage bufferedImage ) {
        try ( final ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream() ) {
            final boolean succeeded = ImageIO.write( bufferedImage, "png", byteArrayOutputStream );
            if ( !succeeded ) {
                return null;
            }

            final byte[] imageByteArray = byteArrayOutputStream.toByteArray();
            try ( final InputStream bais = new ByteArrayInputStream( imageByteArray ) ) {
                final Image image = new Image( document, bais, ImageType.PNG );
                return image;
            }
            catch ( final Exception e ) {
                e.printStackTrace();
                return null;
            }
        }
        catch ( final Exception e ) {
            e.printStackTrace();
            return null;
        }
    }

    // Generic method to get the privacy clause for a PDF report.
    // TODO: Move this to a more general package for reuse in other formats.
    public static String getPrivacyClause() {
        return "This document is for discussion and/or bid purposes only."; //$NON-NLS-1$
    }

    // Generic method to get the report title for a PDF report.
    // TODO: Move this to a more general package for reuse in other formats.
    @SuppressWarnings("nls")
    public static String getReportTitle( final ProductBranding productBranding,
                                         final String reportSubtitle ) {
        final StringBuilder reportTitle = new StringBuilder();
        reportTitle.append( productBranding.productName );
        reportTitle.append( " - " );
        reportTitle.append( reportSubtitle );

        return reportTitle.toString();
    }

    public static void setColumnWidths( final List< Cell > rowData,
                                        final int[] columnWidthsInPixels,
                                        final boolean landscapeMode ) {
        // Determine the scale factor for column widths, which are expressed
        // in pixels but need to be recalculated in points based on the page
        // layout width (which accounts for margins).
        final float columnWidthScaleFactor = getColumnWidthScaleFactor( columnWidthsInPixels,
                                                                        landscapeMode );

        // If manual column widths are present, set column widths now.
        setColumnWidths( rowData, columnWidthsInPixels, columnWidthScaleFactor );
    }

    public static void setColumnWidths( final List< Cell > rowData,
                                        final int[] columnWidthsInPixels,
                                        final float columnWidthScaleFactor ) {
        // If manual column widths are present, set column widths now, after
        // scaling from pixels to points.
        // TODO: Modify the addTableCell method to do this?
        if ( columnWidthsInPixels != null ) {
            for ( int columnWidthIndex =
                                       0; columnWidthIndex < columnWidthsInPixels.length; columnWidthIndex++ ) {
                final float columnWidth = columnWidthsInPixels[ columnWidthIndex ]
                        * columnWidthScaleFactor;
                rowData.get( columnWidthIndex ).setWidth( columnWidth );
            }
        }
    }

    public static void writeFooter( final List< Paragraph > paragraphs,
                                    final PdfFonts fonts,
                                    final ProductBranding productBranding,
                                    final Locale locale ) {
        // Write the name/version of this program and the current user locale as
        // the footer to the first page.
        final String savedFrom = FileUtilities.getSavedFrom( productBranding, locale );
        addParagraph( paragraphs, Align.CENTER, fonts._footerFont, savedFrom );

        // For legal reasons, we also add a privacy clause.
        final String privacyClause = getPrivacyClause();
        addParagraph( paragraphs, Align.CENTER, fonts._footerFont, privacyClause );
    }

    public static void writeFrontPage( final PDF document,
                                       final PdfFonts fonts,
                                       final ProductBranding productBranding,
                                       final String reportSubtitle,
                                       final boolean exportProjectProperties,
                                       final String projectName,
                                       final String venue,
                                       final String designer,
                                       final String date,
                                       final boolean useProjectNotes,
                                       final String projectNotes,
                                       final Locale locale )
            throws Exception {
        // Make a new Page for the Front Page content.
        final Page frontPage = new Page( document, PORTRAIT_PAGE_SIZE );

        // The Front Page is text-only (no tables) so can use a Text Frame.
        final List< Paragraph > paragraphs = new ArrayList<>();

        // Generate the report file title for use within the PDF report.
        final String reportTitle = getReportTitle( productBranding, reportSubtitle );

        // Write the header.
        writeHeader( paragraphs, fonts, reportTitle );

        // Conditionally write the Project Properties as the preface.
        if ( exportProjectProperties ) {
            writeProjectProperties( paragraphs,
                                    fonts,
                                    projectName,
                                    venue,
                                    designer,
                                    date,
                                    useProjectNotes,
                                    projectNotes );
        }

        // Write the footer.
        writeFooter( paragraphs, fonts, productBranding, locale );

        // Make a Text Frame to host the list of Paragraphs.
        final TextFrame textFrame = new TextFrame( paragraphs );

        // We seem to have to manually set our positioning on the page.
        textFrame.setLocation( PORTRAIT_LEFT_MARGIN, PORTRAIT_TOP_MARGIN );
        textFrame.setWidth( PORTRAIT_PAGE_LAYOUT_WIDTH );
        textFrame.setHeight( PORTRAIT_PAGE_LAYOUT_HEIGHT );

        // Add the Text Column to the Front Page.
        textFrame.drawOn( frontPage );
    }

    // Generic method to write the header for a PDF Report.
    public static void writeHeader( final List< Paragraph > paragraphs,
                                    final PdfFonts fonts,
                                    final String reportTitle ) {
        addParagraph( paragraphs, Align.CENTER, fonts._headerFont, reportTitle );
    }

    // Generic method to write a borderless single-column Information Table.
    public static Point writeInformationTable( final PDF document,
                                               final Page page,
                                               final Point initialPoint,
                                               final PdfFonts borderlessTableFonts,
                                               final int align,
                                               final String[] information )
            throws Exception {
        // Information tables only have one column, with multiple rows.
        final List< List< Cell > > informationTableData =
                                                        createInformationTableData( borderlessTableFonts,
                                                                                    align,
                                                                                    information );

        // Write the table to as many pages as are required to fit.
        final Point point = writeInformationTable( document,
                                                   page,
                                                   initialPoint,
                                                   borderlessTableFonts,
                                                   informationTableData );

        return point;
    }

    // Generic method to write a borderless single-column Information Table.
    public static Point writeInformationTable( final PDF document,
                                               final Page page,
                                               final Point initialPoint,
                                               final PdfFonts borderlessTableFonts,
                                               final List< List< Cell > > informationTableData )
            throws Exception {
        // Get a table to use for the Information fields, sans headers.
        final Table informationTable = new Table();

        // Set the table to borderless, to match our on-screen look and feel.
        informationTable.setNoCellBorders();

        // Write the table to as many pages as are required to fit.
        Point point = new Point( initialPoint.getX(), initialPoint.getY() );
        point = writeTable( document,
                            page,
                            point,
                            borderlessTableFonts,
                            informationTableData,
                            informationTable,
                            Table.DATA_HAS_0_HEADER_ROWS,
                            true,
                            false );

        return point;
    }

    // Generic method to export Project Properties to PDF.
    public static void writeProjectProperties( final List< Paragraph > paragraphs,
                                               final PdfFonts fonts,
                                               final String projectName,
                                               final String venue,
                                               final String designer,
                                               final String date,
                                               final boolean useProjectNotes,
                                               final String projectNotes ) {
        // Pad this report sub-section to better set it apart visually.
        addEmptyLines( paragraphs, Align.CENTER, fonts._headerFont, 2 );

        // Write the Project Name as the main header for the Project
        // Properties section.
        final StringBuilder projectHeader = new StringBuilder( "Project: " ); //$NON-NLS-1$
        projectHeader.append( projectName );
        addParagraph( paragraphs,
                      Align.CENTER,
                      fonts._propertiesHeaderFont,
                      projectHeader.toString() );

        // Write the remaining Project Properties as a pseudo-table.
        final StringBuilder venueItem = new StringBuilder( "Venue: " ); //$NON-NLS-1$
        venueItem.append( venue );
        addParagraph( paragraphs, Align.CENTER, fonts._propertiesFont, venueItem.toString() );

        final StringBuilder designerItem = new StringBuilder( "Designer: " ); //$NON-NLS-1$
        designerItem.append( designer );
        addParagraph( paragraphs, Align.CENTER, fonts._propertiesFont, designerItem.toString() );

        final StringBuilder dateItem = new StringBuilder( "Date: " ); //$NON-NLS-1$
        dateItem.append( date );
        addParagraph( paragraphs, Align.CENTER, fonts._propertiesFont, dateItem.toString() );

        // Conditionally write the multi-line Project Notes.
        // TODO: Once up to Java 17, make use of Text Blocks?
        if ( useProjectNotes ) {
            final StringBuilder notesHeader = new StringBuilder( "Project Notes: " ); //$NON-NLS-1$
            addParagraph( paragraphs,
                          Align.CENTER,
                          fonts._notesHeaderFont,
                          notesHeader.toString() );

            // Iterate over each line of the Project Notes.
            final String notesItem = projectNotes;
            final String lines[] = notesItem.split( "\\r?\\n" ); //$NON-NLS-1$
            for ( final String line : lines ) {
                addParagraph( paragraphs, Align.LEFT, fonts._notesFont, line );
            }
        }

        // Pad this report sub-section to better set it apart visually.
        addEmptyLines( paragraphs, Align.CENTER, fonts._headerFont, 2 );
    }

    // Generic method to write a section header for a PDF Report.
    public static void writeSectionHeader( final Page page,
                                           final Point point,
                                           final PdfFonts fonts,
                                           final String sectionTitle ) {
        final TextLine titleLine = new TextLine( fonts._sectionHeaderFont, sectionTitle );

        // Center the section header on the page.
        final float x = 0.5f * ( page.getWidth() - titleLine.getWidth() );
        final float y = point.getY();
        titleLine.setPosition( x, y );

        try {
            titleLine.drawOn( page );
        }
        catch ( final Exception e ) {
            e.printStackTrace();
        }
    }

    // Implementation method to write a multi-page table to a PDF Report.
    public static Point writeTable( final PDF document,
                                    final Page firstPage,
                                    final PdfFonts fonts,
                                    final Table table,
                                    final boolean landscapeMode ) {
        try {
            Page page = firstPage;

            final int numberOfPages = table.getNumberOfPages( page );
            int pageNumber = 1;

            while ( true ) {
                // Draw the table directly on the page as though it is a canvas.
                final Point point = table.drawOn( page );

                if ( !table.hasMoreData() ) {
                    // Allow the table to be drawn again later.
                    table.resetRenderedPagesCount();
                    return point;
                }

                final float[] pageSize = landscapeMode ? LANDSCAPE_PAGE_SIZE : PORTRAIT_PAGE_SIZE;
                page = new Page( document, pageSize );

                // Draw "Page x of N" at the returned point, where "x" =
                // pageNumber and "N" = numberOfPages.
                final String pageCounter = "Page " + pageNumber + " of " //$NON-NLS-1$ //$NON-NLS-2$
                        + numberOfPages;
                page.drawString( fonts._footerFont, pageCounter, point.getX(), point.getY() );

                pageNumber++;
            }
        }
        catch ( final Exception e ) {
            e.printStackTrace();
            return null;
        }
    }

    // Generic method to write multi-page table data to a PDF Report.
    public static Point writeTable( final PDF document,
                                    final Page page,
                                    final Point point,
                                    final PdfFonts fonts,
                                    final List< List< Cell > > tableData,
                                    final Table table,
                                    final int numberOfHeaderRows,
                                    final boolean autoAdjustColumnWidths,
                                    final boolean landscapeMode ) {
        // Set the data on the provided table.
        try {
            table.setData( tableData, numberOfHeaderRows );
        }
        catch ( final Exception e ) {
            e.printStackTrace();
        }

        // Now that the data has been set, auto-adjust columns to fit.
        if ( autoAdjustColumnWidths ) {
            table.autoAdjustColumnWidths();
        }

        // Allow cell text to wrap so that the table doesn't clip on the page.
        // NOTE: There are two versions of this method. This version
        // tokenizes words vs. splitting words up, which of course is more
        // legible and intelligible in the context of a Project Report.
        table.wrapAroundCellText();

        // Set the desired position for the table on the page.
        table.setPosition( point.getX(), point.getY() );

        // Write the table to as many pages as are required to fit.
        return writeTable( document, page, fonts, table, landscapeMode );
    }

    public static double writeVisualization( final PDF document,
                                             final Page visualizationPage,
                                             final PdfFonts fonts,
                                             final String chartLabel,
                                             final double layoutWidth,
                                             final boolean useChart1,
                                             final BufferedImage chart1,
                                             final boolean useChart2,
                                             final BufferedImage chart2,
                                             final boolean useLegend,
                                             final BufferedImage chartLegend )
            throws Exception {
        // We seem to have to manually set our positioning on the page.
        final Point point = new Point( PORTRAIT_LEFT_MARGIN, PORTRAIT_TOP_MARGIN );

        // Write the Chart Label, if it isn't included in the Chart layout.
        if ( chartLabel != null ) {
            writeSectionHeader( visualizationPage, point, fonts, chartLabel );
        }

        // Compute the mapping ratio of the visualization layout to be exported,
        // and use it along with PDF page layout units to determine the scaling
        // ratios.
        // NOTE: We want maximum resolution of the resulting PNG, as people
        // will want to see details when they zoom into the PDF document. The
        // default export is a combination of vector graphics, text, and binary
        // images using an indexed color model without alpha depth. Specifying
        // JPEG causes problems with embedded images such as logos, and can
        // cause a grey-green shadowing of data plots even at 100% quality.
        final double scaleFactor = PORTRAIT_PAGE_LAYOUT_WIDTH / layoutWidth;

        // Scale and output the image sources to fit the PDF page destination.
        final double xOffset = PORTRAIT_LEFT_MARGIN;
        final double yOffset = PORTRAIT_TOP_MARGIN + ( ( chartLabel != null ) ? 30d : 0.0d );
        double chart2AdjustmentY = 0.0d;
        double legendAdjustmentX = 0.0d;
        double metadataAdjustmentY = 0.0d;
        if ( useChart1 ) {
            final Image chart1Image = getImageSnapshot( document, chart1 );
            if ( chart1Image != null ) {
                chart1Image.setPosition( xOffset, yOffset );
                chart1Image.scaleBy( scaleFactor );
                chart1Image.drawOn( visualizationPage );

                // Adjust positioning for the next layout elements.
                chart2AdjustmentY = chart1Image.getHeight();
                legendAdjustmentX = chart1Image.getWidth();
                metadataAdjustmentY += chart1Image.getHeight();
            }
        }
        if ( useChart2 ) {
            final Image chart2Image = getImageSnapshot( document, chart2 );
            if ( chart2Image != null ) {
                chart2Image.setPosition( xOffset, yOffset + chart2AdjustmentY );
                chart2Image.scaleBy( scaleFactor );
                chart2Image.drawOn( visualizationPage );

                // Adjust positioning for the next layout elements.
                legendAdjustmentX = FastMath.max( legendAdjustmentX, chart2Image.getWidth() );
                metadataAdjustmentY += chart2Image.getHeight();
            }
        }
        if ( useLegend ) {
            final Image chartLegendImage = getImageSnapshot( document, chartLegend );
            if ( chartLegendImage != null ) {
                chartLegendImage.setPosition( xOffset + legendAdjustmentX, yOffset );
                chartLegendImage.scaleBy( scaleFactor );
                chartLegendImage.drawOn( visualizationPage );
            }
        }

        return metadataAdjustmentY;
    }

}
