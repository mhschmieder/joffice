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

import com.pdfjet.CoreFont;
import com.pdfjet.Font;
import com.pdfjet.PDF;

/**
 * This is a container class for fonts needed by a PDF Report, using our
 * preferred sizes and styles, etc. The PDFjet library provides a facility for
 * generating embedded fonts, which is what PDF/A compliance requires, so we
 * need a container to hold instances of these fonts as they must be created via
 * a live PDF document object from PDFjet and cannot be made static utility
 * class constants as when we worked with the iText PDF Writer library.
 */
public final class PdfFonts {

    public Font _headerFont;
    public Font _footerFont;
    public Font _sectionHeaderFont;
    public Font _tableHeaderFont;
    public Font _tableCellFont;
    public Font _tableLabelFont;
    public Font _chartLabelFont;
    public Font _axisLabelFont;
    public Font _propertiesHeaderFont;
    public Font _propertiesFont;
    public Font _notesHeaderFont;
    public Font _notesFont;

    public PdfFonts( final PDF document,
                     final boolean needMediumTableFonts,
                     final boolean needSmallTableFonts,
                     final boolean needBorderlessTableFonts ) {
        try {
            // Ensure a consistent header font throughout the report.
            _headerFont = new Font( document, CoreFont.HELVETICA_BOLD );
            _headerFont.setSize( 18f );

            // Ensure a consistent footer font throughout the report.
            _footerFont = new Font( document, CoreFont.HELVETICA );
            _footerFont.setSize( 10f );
            _footerFont.setItalic( true );

            // Ensure a consistent section header font throughout the report.
            _sectionHeaderFont = new Font( document, CoreFont.HELVETICA_BOLD );
            _sectionHeaderFont.setSize( 16f );

            // Ensure a consistent table header font throughout the report.
            _tableHeaderFont = new Font( document, CoreFont.HELVETICA_BOLD );
            _tableHeaderFont.setSize( needMediumTableFonts
                ? 6.5f
                : needSmallTableFonts ? 5f : needBorderlessTableFonts ? 6f : 8f );
            _tableHeaderFont.setItalic( true );

            // Ensure a consistent table cell font throughout the report.
            _tableCellFont = new Font( document, CoreFont.HELVETICA );
            _tableCellFont.setSize( needMediumTableFonts
                ? 6.5f
                : needSmallTableFonts ? 5f : needBorderlessTableFonts ? 6f : 8f );

            // Ensure a consistent table label font throughout the report.
            _tableLabelFont = new Font( document, CoreFont.HELVETICA_BOLD );
            _tableLabelFont.setSize( 10f );

            // Ensure a consistent chart label font throughout the report.
            _chartLabelFont = new Font( document, CoreFont.HELVETICA_BOLD );
            _chartLabelFont.setSize( 15f );

            // Ensure a consistent axis label font throughout the report.
            _axisLabelFont = new Font( document, CoreFont.HELVETICA );
            _axisLabelFont.setSize( 6f );

            // Ensure a consistent properties header font throughout the report.
            _propertiesHeaderFont = new Font( document, CoreFont.HELVETICA_BOLD );
            _propertiesHeaderFont.setSize( 13f );

            // Ensure a consistent properties font throughout the report.
            _propertiesFont = new Font( document, CoreFont.HELVETICA );
            _propertiesFont.setSize( 12f );

            // Ensure a consistent notes header font throughout the report.
            _notesHeaderFont = new Font( document, CoreFont.HELVETICA_BOLD );
            _notesHeaderFont.setSize( 12f );

            // Ensure a consistent notes font throughout the report.
            _notesFont = new Font( document, CoreFont.HELVETICA );
            _notesFont.setSize( 11f );
        }
        catch ( final Exception e ) {
            e.printStackTrace();
        }
    }

}
