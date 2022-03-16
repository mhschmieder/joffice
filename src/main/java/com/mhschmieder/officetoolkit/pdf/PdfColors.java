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

import com.pdfjet.Color;

/**
 * Wrapper class for PDF compatible hex-based colors.
 *
 * TODO: Mimic the PDFjet converter methods in their Color class to invert
 * that shifting algorithm, and apply that to our RGB colors in CSS stylesheets
 * instead, so there is no risk of diverging due to cut/paste and hex vs. RGB.
 */
public final class PdfColors {

    // Define some PDF colors that are custom but get used more than once.
    public static final int BRIGHTYELLOW                                       = 0xFFFF0A;
    public static final int BURGUNDY                                           = 0x3E1010;
    public static final int DARKROYALBLUE                                      = 0x3E5697;
    public static final int DARKSKYBLUE                                        = 0x216FE1;
    public static final int DIMBURGUNDY                                        = 0x621818;
    public static final int GRAY15                                             = 0x272727;
    public static final int GRAY25                                             = 0x404040;
    public static final int GRAY30                                             = 0x4D4D4D;
    public static final int GRAY85                                             = 0xD8D8D8;
    public static final int LIGHTBURGUNDY                                      = 0xDE1010;
    public static final int RUST                                               = 0x621818;

    // Declare the Table Header colors as hex integer references.
    public static final int TABLE_HEADER_BACKGROUND_COLOR                      = DARKROYALBLUE;
    public static final int TABLE_HEADER_TEXT_FILL_COLOR                       = Color.white;

    // Ensure a consistent table label color throughout the report.
    public static final int TABLE_LABEL_BACKGROUND_COLOR                       = Color.pink;
    public static final int TABLE_LABEL_TEXT_FILL_COLOR                        = Color.black;

    // Define polarity "reversed" and "normal" background and foreground colors.
    public static final int POLARITY_NORMAL_BACKGROUND_COLOR                   = GRAY15;
    public static final int POLARITY_NORMAL_FOREGROUND_COLOR                   = Color.white;

    public static final int POLARITY_REVERSED_BACKGROUND_COLOR                 = GRAY85;
    public static final int POLARITY_REVERSED_FOREGROUND_COLOR                 = Color.black;

    // Define mute switch "muted" and "unmuted" background and foreground
    // colors.
    public static final int UNMUTED_BACKGROUND_COLOR                           = DIMBURGUNDY;
    public static final int UNMUTED_FOREGROUND_COLOR                           = Color.white;

    public static final int MUTED_BACKGROUND_COLOR                             = Color.red;
    public static final int MUTED_FOREGROUND_COLOR                             = Color.white;

    // Define PDF color references and/or custom hex values for visible.
    public static final int VISIBLE_BACKGROUND_COLOR                           = BRIGHTYELLOW;
    public static final int VISIBLE_FOREGROUND_COLOR                           = Color.blue;

    // Define PDF color references and/or custom hex values for hidden.
    public static final int HIDDEN_BACKGROUND_COLOR                            = Color.blue;
    public static final int HIDDEN_FOREGROUND_COLOR                            = BRIGHTYELLOW;

    // Define PDF color references and/or custom hex values for enabled.
    public static final int ENABLED_BACKGROUND_COLOR                           = Color.olive;
    public static final int ENABLED_FOREGROUND_COLOR                           = Color.white;

    // Define PDF color references and/or custom hex values for bypassed.
    public static final int BYPASSED_BACKGROUND_COLOR                          = BRIGHTYELLOW;
    public static final int BYPASSED_FOREGROUND_COLOR                          = Color.black;

    // Define PDF color references and/or custom hex values for not applicable.
    public static final int NOT_APPLICABLE_BACKGROUND_COLOR                    =
                                                            BYPASSED_BACKGROUND_COLOR;
    public static final int NOT_APPLICABLE_FOREGROUND_COLOR                    =
                                                            BYPASSED_FOREGROUND_COLOR;

    // Define array groups and array type background and foreground colors.
    public static final int ARRAY_PROCESSING_INACTIVE_BACKGROUND_COLOR         = GRAY30;
    public static final int ARRAY_PROCESSING_INACTIVE_FOREGROUND_COLOR         = Color.white;

    public static final int ARRAY_PROCESSING_ACTIVE_BACKGROUND_COLOR           = DARKSKYBLUE;
    public static final int ARRAY_PROCESSING_ACTIVE_FOREGROUND_COLOR           = Color.white;

    // Define Processing Bypassed and Enabled background and foreground colors.
    public static final int PROCESSING_ENABLED_BACKGROUND_COLOR                =
                                                                ENABLED_BACKGROUND_COLOR;
    public static final int PROCESSING_ENABLED_FOREGROUND_COLOR                =
                                                                ENABLED_FOREGROUND_COLOR;

    public static final int PROCESSING_BYPASSED_BACKGROUND_COLOR               =
                                                                 BYPASSED_BACKGROUND_COLOR;
    public static final int PROCESSING_BYPASSED_FOREGROUND_COLOR               =
                                                                 BYPASSED_FOREGROUND_COLOR;

    // Define Output Channel Unassigned and Output Channel Assigned colors.
    public static final int OUTPUT_CHANNEL_UNASSIGNED_BACKGROUND_COLOR         = GRAY30;
    public static final int OUTPUT_CHANNEL_UNASSIGNED_FOREGROUND_COLOR         = Color.white;

    public static final int OUTPUT_CHANNEL_ASSIGNED_BACKGROUND_COLOR           = DARKSKYBLUE;
    public static final int OUTPUT_CHANNEL_ASSIGNED_FOREGROUND_COLOR           = Color.white;

    // Define Associated Outputs status OK, warning and error colors.
    public static final int ASSOCIATED_OUTPUTS_STATUS_OK_BACKGROUND_COLOR      = Color.white;
    public static final int ASSOCIATED_OUTPUTS_STATUS_OK_FOREGROUND_COLOR      = Color.black;

    public static final int ASSOCIATED_OUTPUTS_STATUS_WARNING_BACKGROUND_COLOR = Color.olive;
    public static final int ASSOCIATED_OUTPUTS_STATUS_WARNING_FOREGROUND_COLOR = Color.white;

    public static final int ASSOCIATED_OUTPUTS_STATUS_ERROR_BACKGROUND_COLOR   = RUST;
    public static final int ASSOCIATED_OUTPUTS_STATUS_ERROR_FOREGROUND_COLOR   = Color.white;

    // Define Pull-Back Load Status colors.
    public static final int PULL_BACK_LOAD_STATUS_OK                           = Color.green;
    public static final int PULL_BACK_LOAD_STATUS_FAIL                         = LIGHTBURGUNDY;

}
