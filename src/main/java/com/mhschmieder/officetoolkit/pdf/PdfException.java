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

/**
 * This is an encapsulation of exceptions thrown by this library, for purposes
 * of clearly identifying library calls as the underlying cause.
 */
public class PdfException extends RuntimeException {
    /**
     *
     */
    private static final long serialVersionUID = 2649797198528955199L;

    /**
     * Fully qualified constructor for a library encapsulation of exceptions
     * related to PDF handling. Generally these will be recaptures of Core Java
     * exceptions, wrapped in a library class to better mark the cause or fault.
     *
     * @param message
     *            The full pre-parsed string to include with the exception
     *
     * @since 1.0
     */
    public PdfException( final String message ) {
        super( message );
    }

}
