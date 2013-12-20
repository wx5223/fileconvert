/*
	This file is part of pdf2html.

    pdf2html is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    pdf2html is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with pdf2html.  If not, see <http://www.gnu.org/licenses/>.
    
    Copyright by timendum.
    Based on org.apache.pdfbox.util.PDFTextStripper by Apache Software Foundation
 */
package net.timendum.pdf;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;

import org.apache.pdfbox.util.TextNormalize;
import org.apache.pdfbox.util.TextPosition;
import org.apache.pdfbox.util.TextPositionComparator;

public class LocalPDFTextStripper extends org.apache.pdfbox.util.PDFText2HTML {

	private static final float ENDOFLASTTEXTX_RESET_VALUE = -1;
	private static final float MAXYFORLINE_RESET_VALUE = -Float.MAX_VALUE;
	private static final float EXPECTEDSTARTOFNEXTWORDX_RESET_VALUE = -Float.MAX_VALUE;
	private static final float MAXHEIGHTFORLINE_RESET_VALUE = -1;
	private static final float MINYTOPFORLINE_RESET_VALUE = Float.MAX_VALUE;
	private static final float LASTWORDSPACING_RESET_VALUE = -1;
	private boolean onFirstPage = true;

	public LocalPDFTextStripper() throws IOException {
		this("UTF-8");
	}

	public LocalPDFTextStripper(String encoding) throws IOException {
		super(encoding);
		normalize = new TextNormalize(outputEncoding);
	}

	@SuppressWarnings("unchecked")
	@Override
	protected void writePage() throws IOException {
		if (onFirstPage) {
			writeHeader();
			onFirstPage = false;
		}
		float maxYForLine = MAXYFORLINE_RESET_VALUE;
		float minYTopForLine = MINYTOPFORLINE_RESET_VALUE;
		float endOfLastTextX = ENDOFLASTTEXTX_RESET_VALUE;
		float lastWordSpacing = LASTWORDSPACING_RESET_VALUE;
		float maxHeightForLine = MAXHEIGHTFORLINE_RESET_VALUE;
		PositionWrapper lastPosition = null;

		boolean startOfPage = true;//flag to indicate start of page
		boolean startOfArticle = true;
		if (charactersByArticle.size() > 0) {
			writePageStart();
		}

		for (int i = 0; i < charactersByArticle.size(); i++) {
			List<TextPosition> textList = charactersByArticle.get(i);
			if (getSortByPosition()) {
				TextPositionComparator comparator = new TextPositionComparator();
				Collections.sort(textList, comparator);
			}

			Iterator<TextPosition> textIter = textList.iterator();

			/* Before we can display the text, we need to do some normalizing.
			 * Arabic and Hebrew text is right to left and is typically stored
			 * in its logical format, which means that the rightmost character is
			 * stored first, followed by the second character from the right etc.
			 * However, PDF stores the text in presentation form, which is left to
			 * right.  We need to do some normalization to convert the PDF data to
			 * the proper logical output format.
			 *
			 * Note that if we did not sort the text, then the output of reversing the
			 * text is undefined and can sometimes produce worse output then not trying
			 * to reverse the order.  Sorting should be done for these languages.
			 * */

			/* First step is to determine if we have any right to left text, and
			 * if so, is it dominant. */
			int ltrCnt = 0;
			int rtlCnt = 0;

			while (textIter.hasNext()) {
				TextPosition position = textIter.next();
				String stringValue = position.getCharacter();
				for (int a = 0; a < stringValue.length(); a++) {
					byte dir = Character.getDirectionality(stringValue.charAt(a));
					if ((dir == Character.DIRECTIONALITY_LEFT_TO_RIGHT)
						|| (dir == Character.DIRECTIONALITY_LEFT_TO_RIGHT_EMBEDDING)
						|| (dir == Character.DIRECTIONALITY_LEFT_TO_RIGHT_OVERRIDE)) {
						ltrCnt++;
					} else if ((dir == Character.DIRECTIONALITY_RIGHT_TO_LEFT)
						|| (dir == Character.DIRECTIONALITY_RIGHT_TO_LEFT_ARABIC)
						|| (dir == Character.DIRECTIONALITY_RIGHT_TO_LEFT_EMBEDDING)
						|| (dir == Character.DIRECTIONALITY_RIGHT_TO_LEFT_OVERRIDE)) {
						rtlCnt++;
					}
				}
			}

			// choose the dominant direction
			boolean isRtlDominant = rtlCnt > ltrCnt;

			startArticle(!isRtlDominant);
			startOfArticle = true;
			// we will later use this to skip reordering
			boolean hasRtl = rtlCnt > 0;

			/* Now cycle through to print the text.
			 * We queue up a line at a time before we print so that we can convert
			 * the line from presentation form to logical form (if needed). 
			 */
			List<TextPosition> line = new ArrayList<TextPosition>();

			textIter = textList.iterator(); // start from the beginning again
			/* PDF files don't always store spaces. We will need to guess where we should add
			 * spaces based on the distances between TextPositions. Historically, this was done
			 * based on the size of the space character provided by the font. In general, this worked
			 * but there were cases where it did not work. Calculating the average character width
			 * and using that as a metric works better in some cases but fails in some cases where the
			 * spacing worked. So we use both. NOTE: Adobe reader also fails on some of these examples.
			 */
			//Keeps track of the previous average character width
			float previousAveCharWidth = -1;
			while (textIter.hasNext()) {
				TextPosition position = textIter.next();
				PositionWrapper current = new PositionWrapper(position);
				String characterValue = position.getCharacter();

				//Resets the average character width when we see a change in font
				// or a change in the font size
				if (lastPosition != null
					&& ((position.getFont() != lastPosition.getTextPosition().getFont()) || (position.getFontSize() != lastPosition
						.getTextPosition().getFontSize()))) {
					previousAveCharWidth = -1;
				}

				float positionX;
				float positionY;
				float positionWidth;
				float positionHeight;

				/* If we are sorting, then we need to use the text direction
				 * adjusted coordinates, because they were used in the sorting. */
				if (getSortByPosition()) {
					positionX = position.getXDirAdj();
					positionY = position.getYDirAdj();
					positionWidth = position.getWidthDirAdj();
					positionHeight = position.getHeightDir();
				} else {
					positionX = position.getX();
					positionY = position.getY();
					positionWidth = position.getWidth();
					positionHeight = position.getHeight();
				}

				//The current amount of characters in a word
				int wordCharCount = position.getIndividualWidths().length;

				/* Estimate the expected width of the space based on the
				 * space character with some margin. */
				float wordSpacing = position.getWidthOfSpace();
				float deltaSpace = 0;
				if ((wordSpacing == 0) || (wordSpacing == Float.NaN)) {
					deltaSpace = Float.MAX_VALUE;
				} else {
					if (lastWordSpacing < 0) {
						deltaSpace = (wordSpacing * getSpacingTolerance());
					} else {
						deltaSpace = (((wordSpacing + lastWordSpacing) / 2f) * getSpacingTolerance());
					}
				}

				/* Estimate the expected width of the space based on the
				 * average character width with some margin. This calculation does not
				 * make a true average (average of averages) but we found that it gave the
				 * best results after numerous experiments. Based on experiments we also found that
				 * .3 worked well. */
				float averageCharWidth = -1;
				if (previousAveCharWidth < 0) {
					averageCharWidth = (positionWidth / wordCharCount);
				} else {
					averageCharWidth = (previousAveCharWidth + (positionWidth / wordCharCount)) / 2f;
				}
				float deltaCharWidth = (averageCharWidth * getAverageCharTolerance());

				//Compares the values obtained by the average method and the wordSpacing method and picks
				//the smaller number.
				float expectedStartOfNextWordX = EXPECTEDSTARTOFNEXTWORDX_RESET_VALUE;
				if (endOfLastTextX != ENDOFLASTTEXTX_RESET_VALUE) {
					if (deltaCharWidth > deltaSpace) {
						expectedStartOfNextWordX = endOfLastTextX + deltaSpace;
					} else {
						expectedStartOfNextWordX = endOfLastTextX + deltaCharWidth;
					}
				}

				if (lastPosition != null) {
					if (startOfArticle) {
						lastPosition.setArticleStart();
						startOfArticle = false;
					}
					// RDD - Here we determine whether this text object is on the current
					// line.  We use the lastBaselineFontSize to handle the superscript
					// case, and the size of the current font to handle the subscript case.
					// Text must overlap with the last rendered baseline text by at least
					// a small amount in order to be considered as being on the same line.

					/* XXX BC: In theory, this check should really check if the next char is in full range
					 * seen in this line. This is what I tried to do with minYTopForLine, but this caused a lot
					 * of regression test failures.  So, I'm leaving it be for now. */
					if (!overlap(positionY, positionHeight, maxYForLine, maxHeightForLine)) {
						writeLine(isRtlDominant, hasRtl, line);
						line.clear();

						writeLineSeparator();

						endOfLastTextX = ENDOFLASTTEXTX_RESET_VALUE;
						expectedStartOfNextWordX = EXPECTEDSTARTOFNEXTWORDX_RESET_VALUE;
						maxYForLine = MAXYFORLINE_RESET_VALUE;
						maxHeightForLine = MAXHEIGHTFORLINE_RESET_VALUE;
						minYTopForLine = MINYTOPFORLINE_RESET_VALUE;
					}

					//Test if our TextPosition starts after a new word would be expected to start.
					if (expectedStartOfNextWordX != EXPECTEDSTARTOFNEXTWORDX_RESET_VALUE
						&& expectedStartOfNextWordX < positionX
						&&
						//only bother adding a space if the last character was not a space
						lastPosition.getTextPosition().getCharacter() != null
						&& !lastPosition.getTextPosition().getCharacter().endsWith(" ")) {
						line.add(WordSeparator.getSeparator());
					}
				}

				if (positionY >= maxYForLine) {
					maxYForLine = positionY;
				}

				// RDD - endX is what PDF considers to be the x coordinate of the
				// end position of the text.  We use it in computing our metrics below.
				endOfLastTextX = positionX + positionWidth;

				// add it to the list
				if (characterValue != null) {
					if (startOfPage && lastPosition == null) {
						writeParagraphStart();//not sure this is correct for RTL?
					}
					line.add(position);
				}
				maxHeightForLine = Math.max(maxHeightForLine, positionHeight);
				minYTopForLine = Math.min(minYTopForLine, positionY - positionHeight);
				lastPosition = current;
				if (startOfPage) {
					lastPosition.setParagraphStart();
					lastPosition.setLineStart();
					startOfPage = false;
				}
				lastWordSpacing = wordSpacing;
				previousAveCharWidth = averageCharWidth;
			}

			// print the final line
			if (line.size() > 0) {
				writeLine(isRtlDominant, hasRtl, line);
			}

			endArticle();
		}
		writePageEnd();
	}

	protected void writeLine(boolean isRtlDominant, boolean hasRtl, List<TextPosition> line) throws IOException {
		if (isLineEmpty(line)) {
			return;
		}
		writeLineStart(line);
		output.flush();
		writeLine(line, isRtlDominant, hasRtl);
		output.flush();
		writeLineEnd(line);
		output.flush();
	}

	protected void writeLine(List<TextPosition> line, boolean isRtlDominant, boolean hasRtl) throws IOException {
		String c, normalized;
		for (TextPosition text : line) {
			if (text instanceof WordSeparator) {
				writeWordSeparator();
			} else {
				c = text.getCharacter();
				normalized = normalize.normalizePres(c);
				writeStringBefore(text, c, normalized);
				writeString(normalized);
				writeStringAfter(text, c, normalized);

			}
		}
	}

	protected void writeStringAfter(TextPosition text, String c, String normalized) throws IOException {
	}

	protected void writeStringBefore(TextPosition text, String c, String normalized) throws IOException {
	}

	protected TextNormalize normalize = null;

	protected void writeLineStart(List<TextPosition> line) throws IOException {

	}

	protected void writeLineEnd(List<TextPosition> line) throws IOException {

	}

	protected static final class WordSeparator extends TextPosition {
		private static final WordSeparator separator = new WordSeparator();

		private WordSeparator() {
		}

		public static final WordSeparator getSeparator() {
			return separator;
		}

	}

	protected boolean overlap(float y1, float height1, float y2, float height2) {
		return within(y1, y2, .1f) || (y2 <= y1 && y2 >= y1 - height1) || (y1 <= y2 && y1 >= y2 - height2);
	}


	protected boolean within(float first, float second, float variance) {
		return second < first + variance && second > first - variance;
	}

	protected String escape(String chars) {
		StringBuilder builder = new StringBuilder(chars.length());
		for (int i = 0; i < chars.length(); ++i) {
			char c = chars.charAt(i);

			if ((c < ' ') || (c > '~')) {
				builder.append("&#").append(c).append(";");
			} else {
				switch (c) {
					case '"':
						builder.append("&quot;");
						break;
					case '&':
						builder.append("&amp;");
						break;
					case '<':
						builder.append("&lt;");
						break;
					case '>':
						builder.append("&gt;");
						break;
					default:
						builder.append(String.valueOf(c));
				}
			}
		}
		return builder.toString();
	}

	protected static TextPosition getFirstTrimmed(List<TextPosition> line) {
		String c;
		for (int i = 0; i < line.size(); i++) {
			if (line.get(i) == null) {
				continue;
			}
			c = line.get(i).getCharacter();
			if (c != null && c.trim().length() > 0) {
				return line.get(i);
			}
		}
		return line.get(0);
	}

	protected static TextPosition getLastTrimmed(List<TextPosition> line) {
		String c;
		for (int i = line.size() - 1; i >= 0; i--) {
			if (line.get(i) == null) {
				continue;
			}
			c = line.get(i).getCharacter();
			if (c != null && c.trim().length() > 0) {
				return line.get(i);
			}
		}
		return line.get(line.size() - 1);
	}

	protected boolean isLineEmpty(List<TextPosition> line) {
		return getFirstTrimmed(line).getCharacter().trim().isEmpty();
	}
}
