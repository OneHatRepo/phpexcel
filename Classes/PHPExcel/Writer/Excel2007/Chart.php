<?php
/**
 * PHPExcel
 *
 * Copyright (c) 2006 - 2014 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel_Writer_Excel2007
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    1.8.0, 2014-03-02
 */


/**
 * PHPExcel_Writer_Excel2007_Chart
 *
 * @category   PHPExcel
 * @package    PHPExcel_Writer_Excel2007
 * @copyright  Copyright (c) 2006 - 2014 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_Writer_Excel2007_Chart extends PHPExcel_Writer_Excel2007_WriterPart
{
	/**
	 * Write charts to XML format
	 *
	 * @param 	PHPExcel_Chart				$pChart
	 * @return 	string 						XML Output
	 * @throws 	PHPExcel_Writer_Exception
	 */
	public function writeChart(PHPExcel_Chart $pChart = null)
	{
		// Create XML writer
		$objWriter = null;
		if ($this->getParentWriter()->getUseDiskCaching()) {
			$objWriter = new PHPExcel_Shared_XMLWriter(PHPExcel_Shared_XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
		} else {
			$objWriter = new PHPExcel_Shared_XMLWriter(PHPExcel_Shared_XMLWriter::STORAGE_MEMORY);
		}
		//	Ensure that data series values are up-to-date before we save
		$pChart->refresh();

		// XML header
		$objWriter->startDocument('1.0','UTF-8','yes');

		// c:chartSpace
		$objWriter->startElement('c:chartSpace');
			$objWriter->writeAttribute('xmlns:c', 'http://schemas.openxmlformats.org/drawingml/2006/chart');
			$objWriter->writeAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
			$objWriter->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');

			$objWriter->startElement('c:date1904');
				$objWriter->writeAttribute('val', 0);
			$objWriter->endElement();
			$objWriter->startElement('c:lang');
				$objWriter->writeAttribute('val', "en-GB");
			$objWriter->endElement();
			$objWriter->startElement('c:roundedCorners');
				$objWriter->writeAttribute('val', 0);
			$objWriter->endElement();

			$this->_writeAlternateContent($objWriter);

			$objWriter->startElement('c:chart');

				$this->_writeTitle($pChart->getTitle(), $objWriter);

				$objWriter->startElement('c:autoTitleDeleted');
					$objWriter->writeAttribute('val', 0);
				$objWriter->endElement();

				$this->_writePlotArea($pChart->getPlotArea(),
									  $pChart->getXAxisLabel(),
									  $pChart->getYAxisLabel(),
									  $objWriter,
									  $pChart->getWorksheet(),
									  $pChart->getXAxisOptions(),
									  $pChart->getYAxisOptions()
									 );

				$this->_writeLegend($pChart->getLegend(), $objWriter);


				$objWriter->startElement('c:plotVisOnly');
					$objWriter->writeAttribute('val', 1);
				$objWriter->endElement();

				$objWriter->startElement('c:dispBlanksAs');
					$objWriter->writeAttribute('val', "gap");
				$objWriter->endElement();

				$objWriter->startElement('c:showDLblsOverMax');
					$objWriter->writeAttribute('val', 0);
				$objWriter->endElement();

			$objWriter->endElement();
			
			// SKOTE MOD - adds a border to the chart
			/*$cBorderColor = "000000";
			$objWriter->startElement('c:spPr');
				$objWriter->startElement('a:ln');
					$objWriter->startElement('a:solidFill');
						$objWriter->startElement('a:srgbClr');
							$objWriter->writeAttribute('val',$cBorderColor);
						$objWriter->endElement();
					$objWriter->endElement();
				$objWriter->endElement();
			 $objWriter->endElement();*/
			// END MOD

			$this->_writePrintSettings($objWriter);

		$objWriter->endElement();

		// Return
		return $objWriter->getData();
	}

	/**
	 * Write Chart Title
	 *
	 * @param	PHPExcel_Chart_Title		$title
	 * @param 	PHPExcel_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @throws 	PHPExcel_Writer_Exception
	 */
	private function _writeTitle(PHPExcel_Chart_Title $title = null, $objWriter)
	{
		if (is_null($title)) {
			return;
		}

		$objWriter->startElement('c:title');
			$objWriter->startElement('c:tx');
				$objWriter->startElement('c:rich');

					$objWriter->startElement('a:bodyPr');
					$objWriter->endElement();

					$objWriter->startElement('a:lstStyle');
					$objWriter->endElement();

					$objWriter->startElement('a:p');

						$caption = $title->getCaption();
						if ((is_array($caption)) && (count($caption) > 0))
							$caption = $caption[0];
						$this->getParentWriter()->getWriterPart('stringtable')->writeRichTextForCharts($objWriter, $caption, 'a');

					$objWriter->endElement();
				$objWriter->endElement();
			$objWriter->endElement();

			$layout = $title->getLayout();
			$this->_writeLayout($layout, $objWriter);

			$objWriter->startElement('c:overlay');
				$objWriter->writeAttribute('val', 0);
			$objWriter->endElement();

		$objWriter->endElement();
	}

	/**
	 * Write Chart Legend
	 *
	 * @param	PHPExcel_Chart_Legend		$legend
	 * @param 	PHPExcel_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @throws 	PHPExcel_Writer_Exception
	 */
	private function _writeLegend(PHPExcel_Chart_Legend $legend = null, $objWriter)
	{
		if (is_null($legend)) {
			return;
		}

		$objWriter->startElement('c:legend');

			$objWriter->startElement('c:legendPos');
				$objWriter->writeAttribute('val', $legend->getPosition());
			$objWriter->endElement();

			$layout = $legend->getLayout();
			$this->_writeLayout($layout, $objWriter);

			$objWriter->startElement('c:overlay');
				$objWriter->writeAttribute('val', ($legend->getOverlay()) ? '1' : '0');
			$objWriter->endElement();

			$objWriter->startElement('c:txPr');
				$objWriter->startElement('a:bodyPr');
				$objWriter->endElement();

				$objWriter->startElement('a:lstStyle');
				$objWriter->endElement();

				$objWriter->startElement('a:p');
					$objWriter->startElement('a:pPr');
						$objWriter->writeAttribute('rtl', 0);

						$objWriter->startElement('a:defRPr');
						$objWriter->endElement();
					$objWriter->endElement();

					$objWriter->startElement('a:endParaRPr');
						$objWriter->writeAttribute('lang', "en-US");
					$objWriter->endElement();

				$objWriter->endElement();
			$objWriter->endElement();

		$objWriter->endElement();
	}

	/**
	 * Write Chart Plot Area
	 *
	 * @param	PHPExcel_Chart_PlotArea		$plotArea
	 * @param	PHPExcel_Chart_Title		$xAxisLabel
	 * @param	PHPExcel_Chart_Title		$yAxisLabel
	 * @param 	PHPExcel_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @throws 	PHPExcel_Writer_Exception
	 */
	private function _writePlotArea(PHPExcel_Chart_PlotArea $plotArea,
									PHPExcel_Chart_Title $xAxisLabel = NULL,
									PHPExcel_Chart_Title $yAxisLabel = NULL,
									$objWriter,
									PHPExcel_Worksheet $pSheet,
									$xAxisOptions,
									$yAxisOptions)
	{
		if (is_null($plotArea)) {
			return;
		}

		$xAxisId = '750913281';
		$yAxisId = '750894082';
		$secondaryXAxisId = '374940792';
		$secondaryYAxisId = '282645664';
		$usesSecondaryAxis = false;
		
		$this->_seriesIndex = 0;
		$objWriter->startElement('c:plotArea');

			$layout = $plotArea->getLayout();

			$this->_writeLayout($layout, $objWriter);

			$chartTypes = self::_getChartType($plotArea);
			$catIsMultiLevelSeries = $valIsMultiLevelSeries = FALSE;
			$plotGroupingType = '';

			foreach($chartTypes as $chartType) {

				$objWriter->startElement('c:'.$chartType);

					$groupCount = $plotArea->getPlotGroupCount();
					for($i = 0; $i < $groupCount; ++$i) {
						$plotGroup = $plotArea->getPlotGroupByIndex($i);
						$groupType = $plotGroup->getPlotType();
						if ($groupType == $chartType) {

							$plotStyle = $plotGroup->getPlotStyle();
							if ($groupType === PHPExcel_Chart_DataSeries::TYPE_RADARCHART) {
								$objWriter->startElement('c:radarStyle');
									$objWriter->writeAttribute('val', $plotStyle );
								$objWriter->endElement();
							} elseif ($groupType === PHPExcel_Chart_DataSeries::TYPE_SCATTERCHART) {
								$objWriter->startElement('c:scatterStyle');
									$objWriter->writeAttribute('val', $plotStyle );
								$objWriter->endElement();
							}

							$this->_writePlotGroup($plotGroup, $chartType, $objWriter, $catIsMultiLevelSeries, $valIsMultiLevelSeries, $plotGroupingType, $pSheet);
							
							if (($chartType !== PHPExcel_Chart_DataSeries::TYPE_PIECHART) &&
								($chartType !== PHPExcel_Chart_DataSeries::TYPE_PIECHART_3D) &&
								($chartType !== PHPExcel_Chart_DataSeries::TYPE_DONUTCHART)) {
		
								$plotGroupOptions = $plotGroup->getPlotOptions();

								if (isset($plotGroupOptions['secondaryAxis']) && $plotGroupOptions['secondaryAxis']) {
									$usesSecondaryAxis = true;
									$objWriter->startElement('c:axId');
										$objWriter->writeAttribute('val', $secondaryXAxisId );
									$objWriter->endElement();
									$objWriter->startElement('c:axId');
										$objWriter->writeAttribute('val', $secondaryYAxisId );
									$objWriter->endElement();
									
								} else {
									$objWriter->startElement('c:axId');
										$objWriter->writeAttribute('val', $xAxisId );
									$objWriter->endElement();
									$objWriter->startElement('c:axId');
										$objWriter->writeAttribute('val', $yAxisId );
									$objWriter->endElement();
								}
								
							} else {
								$objWriter->startElement('c:firstSliceAng');
									$objWriter->writeAttribute('val', 0);
								$objWriter->endElement();
		
								if ($chartType === PHPExcel_Chart_DataSeries::TYPE_DONUTCHART) {
		
									$objWriter->startElement('c:holeSize');
										$objWriter->writeAttribute('val', 50);
									$objWriter->endElement();
								}
							}
						}
					}

					$this->_writeDataLbls($objWriter, $layout);

					if ($chartType === PHPExcel_Chart_DataSeries::TYPE_LINECHART) {
						//	Line only, Line3D can't be smoothed

						$objWriter->startElement('c:smooth');
							$objWriter->writeAttribute('val', (integer) $plotGroup->getSmoothLine() );
						$objWriter->endElement();
						
					} elseif (($chartType === PHPExcel_Chart_DataSeries::TYPE_BARCHART) ||
						($chartType === PHPExcel_Chart_DataSeries::TYPE_BARCHART_3D)) {
						
						$objWriter->startElement('c:gapWidth');
							$objWriter->writeAttribute('val', 36 );
						$objWriter->endElement();
						
						if ($plotGroupingType == 'percentStacked' ||
							$plotGroupingType == 'stacked') {

							$objWriter->startElement('c:overlap');
								$objWriter->writeAttribute('val', 100 );
							$objWriter->endElement();
						}
						
					} elseif ($chartType === PHPExcel_Chart_DataSeries::TYPE_BUBBLECHART) {

							$objWriter->startElement('c:bubbleScale');
								$objWriter->writeAttribute('val', 25 );
							$objWriter->endElement();

							$objWriter->startElement('c:showNegBubbles');
								$objWriter->writeAttribute('val', 0 );
							$objWriter->endElement();
					
					} elseif ($chartType === PHPExcel_Chart_DataSeries::TYPE_STOCKCHART) {

							$objWriter->startElement('c:hiLowLines');
							$objWriter->endElement();

							$objWriter->startElement( 'c:upDownBars' );

							$objWriter->startElement( 'c:gapWidth' );
							$objWriter->writeAttribute('val', 300);
							$objWriter->endElement();

							$objWriter->startElement( 'c:upBars' );
							$objWriter->endElement();

							$objWriter->startElement( 'c:downBars' );
							$objWriter->endElement();

							$objWriter->endElement();
					}

				$objWriter->endElement();
			}

			if (($chartType !== PHPExcel_Chart_DataSeries::TYPE_PIECHART) &&
				($chartType !== PHPExcel_Chart_DataSeries::TYPE_PIECHART_3D) &&
				($chartType !== PHPExcel_Chart_DataSeries::TYPE_DONUTCHART)) {

				if ($chartType === PHPExcel_Chart_DataSeries::TYPE_BUBBLECHART) {
					$this->_writeValAx($objWriter,$plotArea,$xAxisLabel,$chartType,$xAxisId,$yAxisId,$catIsMultiLevelSeries, $xAxisOptions);
				} else {
					$this->_writeCatAx($objWriter,$plotArea,$xAxisLabel,$chartType,$xAxisId,$yAxisId,$catIsMultiLevelSeries, $xAxisOptions);
					if ($usesSecondaryAxis) {
						$options = isset($xAxisOptions['secondary']) ? $xAxisOptions['secondary'] : $xAxisOptions;
						$this->_writeCatAx($objWriter,$plotArea,$xAxisLabel,$chartType,$secondaryXAxisId,$secondaryYAxisId,$catIsMultiLevelSeries, $options);
					}
				}

				$this->_writeValAx($objWriter,$plotArea,$yAxisLabel,$chartType,$xAxisId,$yAxisId,$valIsMultiLevelSeries, $yAxisOptions);
				if ($usesSecondaryAxis) {
					$options = isset($yAxisOptions['secondary']) ? $yAxisOptions['secondary'] : $yAxisOptions;
					$this->_writeValAx($objWriter,$plotArea,$yAxisLabel,$chartType,$secondaryXAxisId,$secondaryYAxisId,$valIsMultiLevelSeries, $options);
				}
			}

		$objWriter->endElement();
	}

	/**
	 * Write Data Labels
	 *
	 * @param 	PHPExcel_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @param 	PHPExcel_Chart_Layout		$chartLayout	Chart layout
	 * @throws 	PHPExcel_Writer_Exception
	 */
	private function _writeDataLbls($objWriter, $chartLayout)
	{
		$objWriter->startElement('c:dLbls');

			$objWriter->startElement('c:showLegendKey');
				$showLegendKey = (empty($chartLayout)) ? 0 : $chartLayout->getShowLegendKey();
				$objWriter->writeAttribute('val', ((empty($showLegendKey)) ? 0 : 1) );
			$objWriter->endElement();


			$objWriter->startElement('c:showVal');
				$showVal = (empty($chartLayout)) ? 0 : $chartLayout->getShowVal();
				$objWriter->writeAttribute('val', ((empty($showVal)) ? 0 : 1) );
			$objWriter->endElement();

			$objWriter->startElement('c:showCatName');
				$showCatName = (empty($chartLayout)) ? 0 : $chartLayout->getShowCatName();
				$objWriter->writeAttribute('val', ((empty($showCatName)) ? 0 : 1) );
			$objWriter->endElement();

			$objWriter->startElement('c:showSerName');
				$showSerName = (empty($chartLayout)) ? 0 : $chartLayout->getShowSerName();
				$objWriter->writeAttribute('val', ((empty($showSerName)) ? 0 : 1) );
			$objWriter->endElement();

			$objWriter->startElement('c:showPercent');
				$showPercent = (empty($chartLayout)) ? 0 : $chartLayout->getShowPercent();
				$objWriter->writeAttribute('val', ((empty($showPercent)) ? 0 : 1) );
			$objWriter->endElement();

			$objWriter->startElement('c:showBubbleSize');
				$showBubbleSize = (empty($chartLayout)) ? 0 : $chartLayout->getShowBubbleSize();
				$objWriter->writeAttribute('val', ((empty($showBubbleSize)) ? 0 : 1) );
			$objWriter->endElement();

			$objWriter->startElement('c:showLeaderLines');
				$showLeaderLines = (empty($chartLayout)) ? 1 : $chartLayout->getShowLeaderLines();
				$objWriter->writeAttribute('val', ((empty($showLeaderLines)) ? 0 : 1) );
			$objWriter->endElement();

		$objWriter->endElement();
	}

	/**
	 * Write Category Axis
	 *
	 * @param 	PHPExcel_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @param 	PHPExcel_Chart_PlotArea		$plotArea
	 * @param 	PHPExcel_Chart_Title		$xAxisLabel
	 * @param 	string						$groupType		Chart type
	 * @param 	string						$axId			ID of this axis
	 * @param 	string						$crossAxId		ID of cross axis
	 * @param 	boolean						$isMultiLevelSeries
	 * @param 	array						$options		styling / formatting options
	 * @throws 	PHPExcel_Writer_Exception
	 */
	private function _writeCatAx($objWriter, PHPExcel_Chart_PlotArea $plotArea, $xAxisLabel, $groupType, $axId, $crossAxId, $isMultiLevelSeries, $options)
	{
		
		// Set option defaults
		// Complete list of options and explanations can be found at:
		// http://www.ecma-international.org/publications/standards/Ecma-376.htm
		// DrawingML / DrawingML-Charts / Elements
				
		// Show this axis? (bool)
		if (!isset($options['delete'])) { $options['delete'] = false; }
		// position of the axis on the chart
		if (!isset($options['axPos'])) { $options['axPos'] = 'b'; } // (l, r, b, t) i.e. Left, Right, Bottom, Top
		// No Multi-level Labels (false = labels are drawn as flat text, true = labels are drawn as a hierarchy)
		$options['noMultiLvlLbl'] = ($isMultiLevelSeries ? false : true); // override user-supplied setting in $options
		
		// Axis Options
			// Fill & Line
				// Fill
					// Fill type (picture, pattern not yet supported)
					// Color
					// Transparency
					// Gradient Type
					// Gradient Direction
					// Gradient Angle
					// Gradient stops
					// Gradient Position
					// Gradient Brightness
				// Line (stroke)
					// Line type
					// Color
					// Transparency
					// Width
					// Compound type
					// Dash type
					// Cap type
					// Join type
					// Begin Arrow type
					// Begin Arrow size
					// End Arrow type
					// End Arrow size
					// Gradient Type
					// Gradient Direction
					// Gradient Angle
					// Gradient stops
					// Gradient Position
					// Gradient Brightness
			// Effects
				// Shadow
					// Color
					// Transparency
					// Size
					// Blur
					// Angle
					// Distance
				// Glow
					// Color
					// Size (string) e.g. "10 pt"
					// Transparency (perc) e.g. "10%"
				// Soft Edges
					// Size (string) e.g. "17 pt"
				// 3-D Format (not supported)
			// Size & Properties
				// Alignment
					// Vertical alignment (Top, Middle, Bottom, Top Centered, Middle Centered, Bottom Centered)
					// Text direction (Horizontal, rotate all text 90deg, Rotate all text 270deg, Stacked)


			// Axis Options
				// Axis Options
					// Axis Type (auto, Text, Date)
					if (!isset($options['auto'])) { $options['auto'] = true; }
					// How does this axis cross the perpendicular axis? (Auto, At category number, At maximum category)
					if (!isset($options['crosses'])) { $options['crosses'] = 'autoZero '; } // autoZero, min, max
					// Vertical axis crosses category number (int)
					if (!isset($options['crossesAt'])) { $options['crossesAt'] = 0; }
					// Axis position (On tick marks , Between tick marks)
					if (!isset($options['crossBetween'])) { $options['crossBetween'] = 'between'; } // (midCat, between). "midCat" can cause the outer parts of a bar chart to become hidden behind the chart margins
					// Category order (Min-Max, Max-Min)
					if (!isset($options['scaling'])) { $options['scaling'] = array(); } // container for other settings
					if (!isset($options['scaling']['orientation'])) { $options['scaling']['orientation'] = 'minMax'; } // (minMax, maxMin)
				// Tick Marks
					// Interval between marks (int)
					if (!isset($options['tickMarkSkip'])) { $options['tickMarkSkip'] = 1; }
					// Major type (None, Inside, Outside, Cross)
					if (!isset($options['majorTickMark'])) { $options['majorTickMark'] = 'out'; } // (none, in, out, cross)
					// Minor type (None, Inside, Outside, Cross)
					if (!isset($options['minorTickMark'])) { $options['minorTickMark'] = 'none'; } // (none, in, out, cross)
				// Labels
					// Alignment (Center, Left, Right)
					if (!isset($options['lblAlgn'])) { $options['lblAlgn'] = 'ctr'; } // (ctr, l, r)
					// Interval between labels (int)
					if (!isset($options['tickLblSkip'])) { $options['tickLblSkip'] = 1; }
					// Distance from axis (int, percentage of offset from parent from 0-1000)
					if (!isset($options['lblOffset'])) { $options['lblOffset'] = 100; }
					// Label position (Next to, High, Low, None)
					if (!isset($options['tickLblPos'])) { $options['tickLblPos'] = 'nextTo'; } // (nextTo, high, low, none)
				// Number
					// Category (General, Number, Currency, Accounting, Date, Time, Percentage, Fraction, Scientific, Text, Special, Custom)
					// Decimal places (int)
					// Use 1000s separator (bool)
					// Symbol (string)
					// Negative numbers (array)
					// Type (string) e.g. date type or fraction type, or zip code type
					// Locale (string)
					// Format Code (string) e.g. "$#,##0.00"
					if (!isset($options['formatCode'])) { $options['formatCode'] = '#,##0_-'; }
					// Linked to source (bool)
					if (!isset($options['sourceLinked'])) { $options['sourceLinked'] = false; }
					
		// Text Options
			// Fill & Outline
				// Fill
					// Fill type (picture, pattern not yet supported)
					// Color
					// Transparency
					// Gradient Type
					// Gradient Direction
					// Gradient Angle
					// Gradient stops
					// Gradient Position
					// Gradient Brightness
				// Outline (stroke)
					// ...
			// Text Effects
			if (!isset($options['text'])) { $options['text'] = array(); } // container for Text Properties
				// Shadow
					// Color
					// Transparency
					// Size
					// Blur
					// Angle
					// Distance
				// Glow
					// Color
					// Size
					// Transparency
				// Soft Edges
					// Size
				// 3-D Format
					// ...
			// Textbox
				// Vertical alignment
				if (!isset($options['text']['vert'])) { $options['text']['vert'] = 'horz'; } // (eaVert, horz, mongolianVert, vert, vert270, wordArtVert, wordArtVertRtl)
				// Text direction / Custom angle (int, representing degrees)
				if (!isset($options['text']['rot'])) { $options['text']['rot'] = 0; }
				// Resize shape to fit text
				// Allow text to overflow shape
				if (!isset($options['text']['horzOverflow'])) { $options['text']['horzOverflow'] = 'overflow'; } // (clip, overflow)
				if (!isset($options['text']['vertOverflow'])) { $options['text']['vertOverflow'] = 'overflow'; } // (clip, ellipsis, overflow)
				// Left margin
				// Right margin
				// Top margin
				// Bottom margin
				// Wrap text in shape
				if (!isset($options['text']['wrap'])) { $options['text']['wrap'] = 'none'; } // (none, square)
								
				// Text to stay upright
				if (!isset($options['text']['upright'])) { $options['text']['upright'] = false; } // (b, ctr, dist, just, t)
				// Anchor
				if (!isset($options['text']['anchor'])) { $options['text']['anchor'] = 'b'; } // (b, ctr, dist, just, t)
				if (!isset($options['text']['anchorCtr'])) { $options['text']['anchorCtr'] = 'b'; } // (b, ctr, dist, just, t)
				
		
		
		// Now write the XML
		$objWriter->startElement('c:catAx');

			if ($axId > 0) {
				$objWriter->startElement('c:axId');
					$objWriter->writeAttribute('val', $axId);
				$objWriter->endElement();
			}

			$objWriter->startElement('c:scaling');
				$objWriter->startElement('c:orientation');
					$objWriter->writeAttribute('val', $options['scaling']['orientation']);
				$objWriter->endElement();
			$objWriter->endElement();

			$objWriter->startElement('c:delete');
				$objWriter->writeAttribute('val', intval($options['delete']));
			$objWriter->endElement();

			$objWriter->startElement('c:axPos');
				$objWriter->writeAttribute('val', $options['axPos']);
			$objWriter->endElement();

			if (!is_null($xAxisLabel)) {
				$objWriter->startElement('c:title');
					$objWriter->startElement('c:tx');
						$objWriter->startElement('c:rich');

							$objWriter->startElement('a:bodyPr');
							$objWriter->endElement();

							$objWriter->startElement('a:lstStyle');
							$objWriter->endElement();

							$objWriter->startElement('a:p');
								$objWriter->startElement('a:r');

									$caption = $xAxisLabel->getCaption();
									if (is_array($caption))
										$caption = $caption[0];
									$objWriter->startElement('a:t');
//										$objWriter->writeAttribute('xml:space', 'preserve');
										$objWriter->writeRawData(PHPExcel_Shared_String::ControlCharacterPHP2OOXML( $caption ));
									$objWriter->endElement();

								$objWriter->endElement();
							$objWriter->endElement();
						$objWriter->endElement();
					$objWriter->endElement();

					$layout = $xAxisLabel->getLayout();
					$this->_writeLayout($layout, $objWriter);

					$objWriter->startElement('c:overlay');
						$objWriter->writeAttribute('val', 0);
					$objWriter->endElement();

				$objWriter->endElement();

			}

			$objWriter->startElement('c:numFmt');
				$objWriter->writeAttribute('formatCode', $options['formatCode']);
				$objWriter->writeAttribute('sourceLinked', intval($options['sourceLinked']));
			$objWriter->endElement();

			if ($options['tickMarkSkip'] > 0) {
				$objWriter->startElement('c:tickMarkSkip');
					$objWriter->writeAttribute('val', $options['tickMarkSkip']);
				$objWriter->endElement();
			}

			$objWriter->startElement('c:majorTickMark');
				$objWriter->writeAttribute('val', $options['majorTickMark']);
			$objWriter->endElement();

			$objWriter->startElement('c:minorTickMark');
				$objWriter->writeAttribute('val', $options['minorTickMark']);
			$objWriter->endElement();

			$objWriter->startElement('c:tickLblPos');
				$objWriter->writeAttribute('val', $options['tickLblPos']);
			$objWriter->endElement();

			if ($options['tickLblSkip'] > 0) {
				$objWriter->startElement('c:tickLblSkip');
					$objWriter->writeAttribute('val', $options['tickLblSkip']);
				$objWriter->endElement();
			}
			
			$objWriter->startElement('c:txPr');
				$objWriter->startElement('a:bodyPr');
					$objWriter->writeAttribute('anchor', $options['text']['anchor']);
					$objWriter->writeAttribute('anchorCtr', intval($options['text']['anchorCtr']));
					$objWriter->writeAttribute('horzOverflow', $options['text']['horzOverflow']);
					$objWriter->writeAttribute('rot', $options['text']['rot'] * 60000);
					$objWriter->writeAttribute('upright', intval($options['text']['upright']));
					$objWriter->writeAttribute('vert', $options['text']['vert']);
					$objWriter->writeAttribute('vertOverflow', $options['text']['vertOverflow']);
					$objWriter->writeAttribute('wrap', $options['text']['wrap']);
				$objWriter->endElement();
				$objWriter->startElement('a:lstStyle');
				$objWriter->endElement();
				$objWriter->startElement('a:p');
					$objWriter->startElement('a:pPr');
						$objWriter->startElement('a:defRPr');
						$objWriter->endElement();
					$objWriter->endElement();
					$objWriter->startElement('a:endParaRPr');
						$objWriter->writeAttribute('lang', 'en-US');
					$objWriter->endElement();
				$objWriter->endElement();
			$objWriter->endElement();

			if ($crossAxId > 0) {
				$objWriter->startElement('c:crossAx');
					$objWriter->writeAttribute('val', $crossAxId);
				$objWriter->endElement();
				
				if ($options['crosses'] === 'auto' || $options['crosses'] === 'max') {
					$objWriter->startElement('c:crosses');
						$objWriter->writeAttribute('val', $options['crosses']);
					$objWriter->endElement();
				} else {
					$objWriter->startElement('c:crosses');
						$objWriter->writeAttribute('val', 'autoZero');
					$objWriter->endElement();
					$objWriter->startElement('c:crossesAt');
						$objWriter->writeAttribute('val', $options['crossesAt']);
					$objWriter->endElement();
				}
				
				$objWriter->startElement('c:crossBetween');
					$objWriter->writeAttribute('val', $options['crossBetween']);
				$objWriter->endElement();
			}

			if ($options['auto']) {
				$objWriter->startElement('c:auto');
					$objWriter->writeAttribute('val', 1);
				$objWriter->endElement();
			}

			$objWriter->startElement('c:lblAlgn');
				$objWriter->writeAttribute('val', $options['lblAlgn']);
			$objWriter->endElement();

			$objWriter->startElement('c:lblOffset');
				$objWriter->writeAttribute('val', $options['lblOffset']);
			$objWriter->endElement();

			$objWriter->startElement('c:noMultiLvlLbl');
				$objWriter->writeAttribute('val', intval($options['noMultiLvlLbl']));
			$objWriter->endElement();

		$objWriter->endElement();

	}


	/**
	 * Write Value Axis
	 *
	 * @param 	PHPExcel_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @param 	PHPExcel_Chart_PlotArea		$plotArea
	 * @param 	PHPExcel_Chart_Title		$yAxisLabel
	 * @param 	string						$groupType		Chart type
	 * @param 	string						$crossAxId		ID of cross axis
	 * @param 	string						$axId			ID of this axis
	 * @param 	boolean						$isMultiLevelSeries
	 * @param 	array						$options		styling / formatting options
	 * @throws 	PHPExcel_Writer_Exception
	 */
	private function _writeValAx($objWriter, PHPExcel_Chart_PlotArea $plotArea, $yAxisLabel, $groupType, $crossAxId, $axId, $isMultiLevelSeries, $options)
	{
		// Set option defaults
		if (!isset($options['majorGridlines'])) { $options['majorGridlines'] = true; }
		
		// Set option defaults
		// Complete list of options and explanations can be found at:
		// http://www.ecma-international.org/publications/standards/Ecma-376.htm
		// DrawingML / DrawingML-Charts / Elements
				
		// Show this axis? (bool)
		if (!isset($options['delete'])) { $options['delete'] = false; }
		// position of the axis on the chart
		if (!isset($options['axPos'])) { $options['axPos'] = 'l'; } // (l, r, b, t) i.e. Left, Right, Bottom, Top
		// No Multi-level Labels (false = labels are drawn as flat text, true = labels are drawn as a hierarchy)
		$options['noMultiLvlLbl'] = ($isMultiLevelSeries ? false : true); // override user-supplied setting in $options
				
		// Axis Options
			// Fill & Line
				// Fill
					// Fill type (picture, pattern not yet supported)
					// Color
					// Transparency
					// Gradient Type
					// Gradient Direction
					// Gradient Angle
					// Gradient stops
					// Gradient Position
					// Gradient Brightness
				// Line (stroke)
					// Line type
					// Color
					// Transparency
					// Width
					// Compound type
					// Dash type
					// Cap type
					// Join type
					// Begin Arrow type
					// Begin Arrow size
					// End Arrow type
					// End Arrow size
					// Gradient Type
					// Gradient Direction
					// Gradient Angle
					// Gradient stops
					// Gradient Position
					// Gradient Brightness
			// Effects
				// Shadow
					// Color
					// Transparency
					// Size
					// Blur
					// Angle
					// Distance
				// Glow
					// Color
					// Size (string) e.g. "10 pt"
					// Transparency (perc) e.g. "10%"
				// Soft Edges
					// Size (string) e.g. "17 pt"
				// 3-D Format (not supported)
			// Size & Properties
				// Alignment
					// Vertical alignment (Top, Middle, Bottom, Top Centered, Middle Centered, Bottom Centered)
					// Text direction (Horizontal, rotate all text 90deg, Rotate all text 270deg, Stacked)
			
			
			// Axis Options
				// Axis Options
					// Bounds min (float)
					if (!isset($options['min'])) { $options['min'] = null; }
					// Bounds max (float)
					if (!isset($options['max'])) { $options['max'] = null; }
					// Units Major (float)
					if (!isset($options['majorUnit'])) { $options['majorUnit'] = null; }
					// Units Minor (float)
					if (!isset($options['minorUnit'])) { $options['minorUnit'] = null; }
					// How does this axis cross the perpendicular axis? (Auto, At category number, At maximum category)
					if (!isset($options['crosses'])) { $options['crosses'] = 'autoZero '; } // autoZero, min, max
					// Vertical axis crosses category number (float)
					if (!isset($options['crossesAt'])) { $options['crossesAt'] = 0; }
					// Axis position (On tick marks , Between tick marks)
					if (!isset($options['crossBetween'])) { $options['crossBetween'] = 'between'; } // (midCat, between). "midCat" can cause the outer parts of a bar chart to become hidden behind the chart margins
					// Value order (Min-Max, Max-Min)
					if (!isset($options['scaling'])) { $options['scaling'] = array(); } // container for other settings
					if (!isset($options['scaling']['orientation'])) { $options['scaling']['orientation'] = 'minMax'; } // (minMax, maxMin)
					
					// Display units (None, Hundreds, Thousands, 10000, 100000, Millions, 10000000, 100000000, Billions, Trillions)
					// Show display units label on chart (bool)
					// Logarithmic scale (bool)
					// Logarithmic scale base (int)
					if (!isset($options['logBase'])) { $options['logBase'] = 10; } // (2-1000)
					// Values in reverse order (bool)
				// Tick Marks
					// Interval between marks (int)
					if (!isset($options['tickMarkSkip'])) { $options['tickMarkSkip'] = 1; }
					// Major type (None, Inside, Outside, Cross)
					if (!isset($options['majorTickMark'])) { $options['majorTickMark'] = 'out'; } // (none, in, out, cross)
					// Minor type (None, Inside, Outside, Cross)
					if (!isset($options['minorTickMark'])) { $options['minorTickMark'] = 'none'; } // (none, in, out, cross)
				// Labels
					// Alignment (Center, Left, Right)
					if (!isset($options['lblAlgn'])) { $options['lblAlgn'] = 'ctr'; } // (ctr, l, r)
					// Interval between labels (int)
					if (!isset($options['tickLblSkip'])) { $options['tickLblSkip'] = 1; }
					// Distance from axis (int, percentage of offset from parent from 0-1000)
					if (!isset($options['lblOffset'])) { $options['lblOffset'] = 100; }
					// Label position (Next to, High, Low, None)
					if (!isset($options['tickLblPos'])) { $options['tickLblPos'] = 'nextTo'; } // (nextTo, high, low, none)
				// Number
					// Category (General, Number, Currency, Accounting, Date, Time, Percentage, Fraction, Scientific, Text, Special, Custom)
					// Decimal places (int)
					// Use 1000s separator (bool)
					// Symbol (string)
					// Negative numbers (array)
					// Type (string) e.g. date type or fraction type, or zip code type
					// Locale (string)
					// Format Code (string) e.g. "$#,##0.00"
					if (!isset($options['formatCode'])) { $options['formatCode'] = '#,##0_-'; }
					// Linked to source (bool)
					if (!isset($options['sourceLinked'])) { $options['sourceLinked'] = false; }
		// Text Options
			// Fill & Outline
				// Fill
					// Fill type (picture, pattern not yet supported)
					// Color
					// Transparency
					// Gradient Type
					// Gradient Direction
					// Gradient Angle
					// Gradient stops
					// Gradient Position
					// Gradient Brightness
				// Outline (stroke) (not yet supported)
			// Text Effects
			if (!isset($options['text'])) { $options['text'] = array(); } // container for Text Properties
				// Shadow
					// Color
					// Transparency
					// Size
					// Blur
					// Angle
					// Distance
				// Glow
					// Color
					// Size
					// Transparency
				// Soft Edges
					// Size
				// 3-D Format (not supported)
			// Textbox
				// Vertical alignment
				if (!isset($options['text']['vert'])) { $options['text']['vert'] = 'horz'; } // (eaVert, horz, mongolianVert, vert, vert270, wordArtVert, wordArtVertRtl)
				// Text direction / Custom angle (int, representing degrees)
				if (!isset($options['text']['rot'])) { $options['text']['rot'] = 0; }
				// Resize shape to fit text
				// Allow text to overflow shape
				if (!isset($options['text']['horzOverflow'])) { $options['text']['horzOverflow'] = 'overflow'; } // (clip, overflow)
				if (!isset($options['text']['vertOverflow'])) { $options['text']['vertOverflow'] = 'overflow'; } // (clip, ellipsis, overflow)
				// Left margin
				// Right margin
				// Top margin
				// Bottom margin
				// Wrap text in shape
				if (!isset($options['text']['wrap'])) { $options['text']['wrap'] = 'none'; } // (none, square)
								
				// Text to stay upright
				if (!isset($options['text']['upright'])) { $options['text']['upright'] = false; } // (b, ctr, dist, just, t)
				// Anchor
				if (!isset($options['text']['anchor'])) { $options['text']['anchor'] = 'b'; } // (b, ctr, dist, just, t)
				if (!isset($options['text']['anchorCtr'])) { $options['text']['anchorCtr'] = 'b'; } // (b, ctr, dist, just, t)
			

				
		
		// Now actually build XML
		$objWriter->startElement('c:valAx');

			if ($axId > 0) {
				$objWriter->startElement('c:axId');
					$objWriter->writeAttribute('val', $axId);
				$objWriter->endElement();
			}

			$objWriter->startElement('c:scaling');
				$objWriter->startElement('c:orientation');
					$objWriter->writeAttribute('val', $options['scaling']['orientation']);
				$objWriter->endElement();
				if (isset($options['min'])) {
					$objWriter->startElement('c:min');
						$objWriter->writeAttribute('val', $options['min']);
					$objWriter->endElement();
				}
				if (isset($options['max'])) {
					$objWriter->startElement('c:max');
						$objWriter->writeAttribute('val', $options['max']);
					$objWriter->endElement();
				}
			$objWriter->endElement();

			$objWriter->startElement('c:delete');
				$objWriter->writeAttribute('val', intval($options['delete']));
			$objWriter->endElement();

			$objWriter->startElement('c:axPos');
				$objWriter->writeAttribute('val', $options['axPos']);
			$objWriter->endElement();

			$objWriter->startElement('c:majorGridlines');
			$objWriter->endElement();
			

			if (!is_null($yAxisLabel)) {
				$objWriter->startElement('c:title');
					$objWriter->startElement('c:tx');
						$objWriter->startElement('c:rich');

							$objWriter->startElement('a:bodyPr');
							$objWriter->endElement();

							$objWriter->startElement('a:lstStyle');
							$objWriter->endElement();

							$objWriter->startElement('a:p');
								$objWriter->startElement('a:r');

									$caption = $yAxisLabel->getCaption();
									if (is_array($caption))
										$caption = $caption[0];
									$objWriter->startElement('a:t');
//										$objWriter->writeAttribute('xml:space', 'preserve');
										$objWriter->writeRawData(PHPExcel_Shared_String::ControlCharacterPHP2OOXML( $caption ));
									$objWriter->endElement();

								$objWriter->endElement();
							$objWriter->endElement();
						$objWriter->endElement();
					$objWriter->endElement();

					if ($groupType !== PHPExcel_Chart_DataSeries::TYPE_BUBBLECHART) {
						$layout = $yAxisLabel->getLayout();
						$this->_writeLayout($layout, $objWriter);
					}

					$objWriter->startElement('c:overlay');
						$objWriter->writeAttribute('val', 0);
					$objWriter->endElement();

				$objWriter->endElement();
			}

			$objWriter->startElement('c:numFmt');
				$objWriter->writeAttribute('formatCode', $options['formatCode']);
				$objWriter->writeAttribute('sourceLinked', intval($options['sourceLinked']));
			$objWriter->endElement();

			if ($options['tickMarkSkip'] > 0) {
				$objWriter->startElement('c:tickMarkSkip');
					$objWriter->writeAttribute('val', $options['tickMarkSkip']);
				$objWriter->endElement();
			}
			
			$objWriter->startElement('c:majorTickMark');
				$objWriter->writeAttribute('val', $options['majorTickMark']);
			$objWriter->endElement();

			$objWriter->startElement('c:minorTickMark');
				$objWriter->writeAttribute('val', $options['minorTickMark']);
			$objWriter->endElement();
			
			$objWriter->startElement('c:tickLblPos');
				$objWriter->writeAttribute('val', $options['tickLblPos']);
			$objWriter->endElement();

			if ($options['tickLblSkip'] > 0) {
				$objWriter->startElement('c:tickLblSkip');
					$objWriter->writeAttribute('val', $options['tickLblSkip']);
				$objWriter->endElement();
			}
			
			$objWriter->startElement('c:txPr');
				$objWriter->startElement('a:bodyPr');
					$objWriter->writeAttribute('anchor', $options['text']['anchor']);
					$objWriter->writeAttribute('anchorCtr', intval($options['text']['anchorCtr']));
					$objWriter->writeAttribute('horzOverflow', $options['text']['horzOverflow']);
					$objWriter->writeAttribute('rot', $options['text']['rot'] * 60000);
					$objWriter->writeAttribute('upright', intval($options['text']['upright']));
					$objWriter->writeAttribute('vert', $options['text']['vert']);
					$objWriter->writeAttribute('vertOverflow', $options['text']['vertOverflow']);
					$objWriter->writeAttribute('wrap', $options['text']['wrap']);
				$objWriter->endElement();
				$objWriter->startElement('a:lstStyle');
				$objWriter->endElement();
				$objWriter->startElement('a:p');
					$objWriter->startElement('a:pPr');
						$objWriter->startElement('a:defRPr');
						$objWriter->endElement();
					$objWriter->endElement();
					$objWriter->startElement('a:endParaRPr');
						$objWriter->writeAttribute('lang', 'en-US');
					$objWriter->endElement();
				$objWriter->endElement();
			$objWriter->endElement();
			
			
			if (isset($options['majorUnit'])) {
				$objWriter->startElement('c:majorUnit');
					$objWriter->writeAttribute('val', $options['majorUnit']);
				$objWriter->endElement();
			}
			if (isset($options['minorUnit'])) {
				$objWriter->startElement('c:minorUnit');
					$objWriter->writeAttribute('val', $options['minorUnit']);
				$objWriter->endElement();
			}

			if ($crossAxId > 0) {
				$objWriter->startElement('c:crossAx');
					$objWriter->writeAttribute('val', $crossAxId);
				$objWriter->endElement();

				if ($options['crosses'] === 'auto' || $options['crosses'] === 'max') {
					$objWriter->startElement('c:crosses');
						$objWriter->writeAttribute('val', $options['crosses']);
					$objWriter->endElement();
				} else {
					$objWriter->startElement('c:crosses');
						$objWriter->writeAttribute('val', 'autoZero');
					$objWriter->endElement();
					$objWriter->startElement('c:crossesAt');
						$objWriter->writeAttribute('val', $options['crossesAt']);
					$objWriter->endElement();
				}
				
				$objWriter->startElement('c:crossBetween');
					$objWriter->writeAttribute('val', $options['crossBetween']);
				$objWriter->endElement();
			}

			$objWriter->startElement('c:lblAlgn');
				$objWriter->writeAttribute('val', $options['lblAlgn']);
			$objWriter->endElement();

			$objWriter->startElement('c:lblOffset');
				$objWriter->writeAttribute('val', $options['lblOffset']);
			$objWriter->endElement();

			if ($groupType !== PHPExcel_Chart_DataSeries::TYPE_BUBBLECHART) {
				$objWriter->startElement('c:noMultiLvlLbl');
					$objWriter->writeAttribute('val', intval($options['noMultiLvlLbl']));
				$objWriter->endElement();
			}
		$objWriter->endElement();

	}


	/**
	 * Get the data series type(s) for a chart plot series
	 *
	 * @param 	PHPExcel_Chart_PlotArea		$plotArea
	 * @return	string|array
	 * @throws 	PHPExcel_Writer_Exception
	 */
	private static function _getChartType($plotArea)
	{
		$groupCount = $plotArea->getPlotGroupCount();

		if ($groupCount == 1) {
			$chartType = array($plotArea->getPlotGroupByIndex(0)->getPlotType());
		} else {
			$chartTypes = array();
			for($i = 0; $i < $groupCount; ++$i) {
				$chartTypes[] = $plotArea->getPlotGroupByIndex($i)->getPlotType();
			}
			$chartType = array_unique($chartTypes);
			if (count($chartTypes) == 0) {
				throw new PHPExcel_Writer_Exception('Chart is not yet implemented');
			}
		}

		return $chartType;
	}

	/**
	 * Write Plot Group (series of related plots)
	 *
	 * @param	PHPExcel_Chart_DataSeries		$plotGroup
	 * @param	string							$groupType				Type of plot for dataseries
	 * @param 	PHPExcel_Shared_XMLWriter 		$objWriter 				XML Writer
	 * @param	boolean							&$catIsMultiLevelSeries	Is category a multi-series category
	 * @param	boolean							&$valIsMultiLevelSeries	Is value set a multi-series set
	 * @param	string							&$plotGroupingType		Type of grouping for multi-series values
	 * @param	PHPExcel_Worksheet 				$pSheet
	 * @throws 	PHPExcel_Writer_Exception
	 */
	private function _writePlotGroup( $plotGroup,
									  $groupType,
									  $objWriter,
									  &$catIsMultiLevelSeries,
									  &$valIsMultiLevelSeries,
									  &$plotGroupingType,
									  PHPExcel_Worksheet $pSheet
									)
	{
		if (is_null($plotGroup)) {
			return;
		}
		
		// Set defaults
		$options = $plotGroup->getPlotOptions();
		$presetColorSchemes = array(
			'bold' => array( 
				array( 'line' => array('srgb' => 'a54a03', 'w' => 12700), 'fill' => array('srgb' => 'ff7800', 'alpha' => 100000, ), ),
				array( 'line' => array('srgb' => '011d68', 'w' => 12700), 'fill' => array('srgb' => '123cac', 'alpha' => 100000, ), ),
				array( 'line' => array('srgb' => 'ccab00', 'w' => 12700), 'fill' => array('srgb' => 'ffd600', 'alpha' => 100000, ), ),
				array( 'line' => array('srgb' => '00991d', 'w' => 12700), 'fill' => array('srgb' => '00bb23', 'alpha' => 100000, ), ),
				array( 'line' => array('srgb' => '38026a', 'w' => 12700), 'fill' => array('srgb' => '6212ac', 'alpha' => 100000, ), ),
				array( 'line' => array('srgb' => 'a60033', 'w' => 12700), 'fill' => array('srgb' => 'd90042', 'alpha' => 100000, ), ),
				array( 'line' => array('srgb' => '0091b0', 'w' => 12700), 'fill' => array('srgb' => '00b3d9', 'alpha' => 100000, ), ),
			),
			'warmCool' => array(
				array( 'line' => array('srgb' => 'e5bb14', 'w' => 12700), 'fill' => array('srgb' => 'ffd738', 'alpha' => 100000, ), ),
				array( 'line' => array('srgb' => 'a54a03', 'w' => 12700), 'fill' => array('srgb' => 'af5916', 'alpha' => 100000, ), ),
				array( 'line' => array('srgb' => 'bc6600', 'w' => 12700), 'fill' => array('srgb' => 'fa8800', 'alpha' => 100000, ), ),
				array( 'line' => array('srgb' => '2cc1ff', 'w' => 12700), 'fill' => array('srgb' => '56c9f9', 'alpha' => 100000, ), ),
				array( 'line' => array('srgb' => '002a99', 'w' => 12700), 'fill' => array('srgb' => '123cac', 'alpha' => 100000, ), ),
				array( 'line' => array('srgb' => '1351f6', 'w' => 12700), 'fill' => array('srgb' => '376dfe', 'alpha' => 100000, ), ),
				array( 'line' => array('srgb' => '0091b0', 'w' => 12700), 'fill' => array('srgb' => '00b3d9', 'alpha' => 100000, ), ),
			),
		);
		$presetGradientSchemes = array(
			'bold' => array(
				array( 'line' => array('srgb' => 'a54a03', 'w' => 12700), 'ang' => 5400000, 'fill' => array( array( 'pos' => 0, 'srgb' => 'ffa759', 'alpha' => 100000, ), array( 'pos' => 47000, 'srgb' => 'ff7800', 'alpha' => 100000, ), array( 'pos' => 100000, 'srgb' => '994902', 'alpha' => 100000, ), ),),
				array( 'line' => array('srgb' => '011d68', 'w' => 12700), 'ang' => 5400000, 'fill' => array( array( 'pos' => 0, 'srgb' => '5077de', 'alpha' => 100000, ), array( 'pos' => 47000, 'srgb' => '123cac', 'alpha' => 100000, ), array( 'pos' => 100000, 'srgb' => '01164d', 'alpha' => 100000, ), ),),
				array( 'line' => array('srgb' => 'ccab00', 'w' => 12700), 'ang' => 5400000, 'fill' => array( array( 'pos' => 0, 'srgb' => 'ffea7e', 'alpha' => 100000, ), array( 'pos' => 47000, 'srgb' => 'ffd600', 'alpha' => 100000, ), array( 'pos' => 100000, 'srgb' => 'd9b600', 'alpha' => 100000, ), ),),
				array( 'line' => array('srgb' => '00991d', 'w' => 12700), 'ang' => 5400000, 'fill' => array( array( 'pos' => 0, 'srgb' => '4bf86b', 'alpha' => 100000, ), array( 'pos' => 47000, 'srgb' => '00bb23', 'alpha' => 100000, ), array( 'pos' => 100000, 'srgb' => '006d15', 'alpha' => 100000, ), ),),
				array( 'line' => array('srgb' => '38026a', 'w' => 12700), 'ang' => 5400000, 'fill' => array( array( 'pos' => 0, 'srgb' => '9749df', 'alpha' => 100000, ), array( 'pos' => 47000, 'srgb' => '6212ac', 'alpha' => 100000, ), array( 'pos' => 100000, 'srgb' => '2f0159', 'alpha' => 100000, ), ),),
				array( 'line' => array('srgb' => 'a60033', 'w' => 12700), 'ang' => 5400000, 'fill' => array( array( 'pos' => 0, 'srgb' => 'fb4e82', 'alpha' => 100000, ), array( 'pos' => 47000, 'srgb' => 'd90042', 'alpha' => 100000, ), array( 'pos' => 100000, 'srgb' => '91002c', 'alpha' => 100000, ), ),),
				array( 'line' => array('srgb' => '0091b0', 'w' => 12700), 'ang' => 5400000, 'fill' => array( array( 'pos' => 0, 'srgb' => '6fe6ff', 'alpha' => 100000, ), array( 'pos' => 47000, 'srgb' => '00b3d9', 'alpha' => 100000, ), array( 'pos' => 100000, 'srgb' => '00829e', 'alpha' => 100000, ), ),),
			),
			'warmCool' => array(
				array( 'line' => array('srgb' => 'e5bb14', 'w' => 12700), 'ang' => 5400000, 'fill' => array( array( 'pos' => 0, 'srgb' => 'fff0b5', 'alpha' => 100000, ), array( 'pos' => 47000, 'srgb' => 'ffd738', 'alpha' => 100000, ), array( 'pos' => 100000, 'srgb' => 'c8a416', 'alpha' => 100000, ), ),),
				array( 'line' => array('srgb' => 'a54a03', 'w' => 12700), 'ang' => 5400000, 'fill' => array( array( 'pos' => 0, 'srgb' => 'e38943', 'alpha' => 100000, ), array( 'pos' => 47000, 'srgb' => 'af5916', 'alpha' => 100000, ), array( 'pos' => 100000, 'srgb' => '6b3104', 'alpha' => 100000, ),),),
				array( 'line' => array('srgb' => 'bc6600', 'w' => 12700), 'ang' => 5400000, 'fill' => array( array( 'pos' => 0, 'srgb' => 'faa846', 'alpha' => 100000, ), array( 'pos' => 47000, 'srgb' => 'fa8800', 'alpha' => 100000, ), array( 'pos' => 100000, 'srgb' => 'a55b02', 'alpha' => 100000, ),),),
				array( 'line' => array('srgb' => '2cc1ff', 'w' => 12700), 'ang' => 5400000, 'fill' => array( array( 'pos' => 0, 'srgb' => 'b5e9ff', 'alpha' => 100000, ), array( 'pos' => 47000, 'srgb' => '56c9f9', 'alpha' => 100000, ), array( 'pos' => 100000, 'srgb' => '077eb0', 'alpha' => 100000, ),),),
				array( 'line' => array('srgb' => '002a99', 'w' => 12700), 'ang' => 5400000, 'fill' => array( array( 'pos' => 0, 'srgb' => '3568ee', 'alpha' => 100000, ), array( 'pos' => 47000, 'srgb' => '123cac', 'alpha' => 100000, ), array( 'pos' => 100000, 'srgb' => '031b5b', 'alpha' => 100000, ),),),
				array( 'line' => array('srgb' => '1351f6', 'w' => 12700), 'ang' => 5400000, 'fill' => array( array( 'pos' => 0, 'srgb' => '8fadff', 'alpha' => 100000, ), array( 'pos' => 47000, 'srgb' => '376dfe', 'alpha' => 100000, ), array( 'pos' => 100000, 'srgb' => '063dd1', 'alpha' => 100000, ),),),
				array( 'line' => array('srgb' => '0091b0', 'w' => 12700), 'ang' => 5400000, 'fill' => array( array( 'pos' => 0, 'srgb' => '6fe6ff', 'alpha' => 100000, ), array( 'pos' => 47000, 'srgb' => '00b3d9', 'alpha' => 100000, ), array( 'pos' => 100000, 'srgb' => '00829e', 'alpha' => 100000, ), ),),
			),
		);
		$colorIx = 0;
		if (isset($options['colors']) && is_string($options['colors'])) {
			$options['colors'] = $presetColorSchemes[ $options['colors'] ]; // user can specify an array of colors, or a string name of one of the preset arrays of colors
		}
		if (isset($options['gradients']) && is_string($options['gradients'])) {
			$options['gradients'] = $presetGradientSchemes[ $options['gradients'] ]; // user can specify an array of colors, or a string name of one of the preset arrays of colors
		}
		if (isset($options['showTrendline']) && $options['showTrendline']) {
			if (!isset($options['trendline'])) { $options['trendline'] = array(); }
			if (!isset($options['trendline']['name'])) { $options['trendline']['name'] = 'Trend'; }
			if (!isset($options['trendline']['trendlineType'])) { $options['trendline']['trendlineType'] = PHPExcel_Chart_DataSeries::TRENDLINE_LINEAR; }
		//	if (!isset($options['trendline']['intercept'])) { $options['trendline']['intercept'] = 0; }
			if (!isset($options['trendline']['dispRSqr'])) { $options['trendline']['dispRSqr'] = 0; }
			if (!isset($options['trendline']['dispEq'])) { $options['trendline']['dispEq'] = 0; }
		}
	
		
		
		// Now start building the XML
		if (($groupType == PHPExcel_Chart_DataSeries::TYPE_BARCHART) ||
			($groupType == PHPExcel_Chart_DataSeries::TYPE_BARCHART_3D)) {
			$objWriter->startElement('c:barDir');
				$objWriter->writeAttribute('val', $plotGroup->getPlotDirection());
			$objWriter->endElement();
		}

		if (!is_null($plotGroup->getPlotGrouping())) {
			$plotGroupingType = $plotGroup->getPlotGrouping();
			$objWriter->startElement('c:grouping');
				$objWriter->writeAttribute('val', $plotGroupingType);
			$objWriter->endElement();
		}

		//	Get these details before the loop, because we can use the count to check for varyColors
		$plotSeriesOrder = $plotGroup->getPlotOrder();
		$plotSeriesCount = count($plotSeriesOrder);

		if (($groupType !== PHPExcel_Chart_DataSeries::TYPE_RADARCHART) &&
			($groupType !== PHPExcel_Chart_DataSeries::TYPE_STOCKCHART)) {

			if ($groupType !== PHPExcel_Chart_DataSeries::TYPE_LINECHART) {
				if (($groupType == PHPExcel_Chart_DataSeries::TYPE_PIECHART) ||
					($groupType == PHPExcel_Chart_DataSeries::TYPE_PIECHART_3D) ||
					($groupType == PHPExcel_Chart_DataSeries::TYPE_DONUTCHART) ||
					($plotSeriesCount > 1)) {
					$objWriter->startElement('c:varyColors');
						$objWriter->writeAttribute('val', 1);
					$objWriter->endElement();
				} else {
					$objWriter->startElement('c:varyColors');
						$objWriter->writeAttribute('val', 0);
					$objWriter->endElement();
				}
			}
		}
		

		foreach($plotSeriesOrder as $plotSeriesIdx => $plotSeriesRef) {
			
			$objWriter->startElement('c:ser');

				$objWriter->startElement('c:idx');
					$objWriter->writeAttribute('val', $this->_seriesIndex + $plotSeriesIdx);
				$objWriter->endElement();

				$objWriter->startElement('c:order');
					$objWriter->writeAttribute('val', $this->_seriesIndex + $plotSeriesRef);
				$objWriter->endElement();

				if (($groupType == PHPExcel_Chart_DataSeries::TYPE_PIECHART) ||
					($groupType == PHPExcel_Chart_DataSeries::TYPE_PIECHART_3D) ||
					($groupType == PHPExcel_Chart_DataSeries::TYPE_DONUTCHART)) {

					$objWriter->startElement('c:dPt');
						$objWriter->startElement('c:idx');
							$objWriter->writeAttribute('val', 3);
						$objWriter->endElement();

						$objWriter->startElement('c:bubble3D');
							$objWriter->writeAttribute('val', 0);
						$objWriter->endElement();

						$objWriter->startElement('c:spPr');
							$objWriter->startElement('a:solidFill');
								$objWriter->startElement('a:srgbClr');
									$objWriter->writeAttribute('val', 'FF9900');
								$objWriter->endElement();
							$objWriter->endElement();
						$objWriter->endElement();
					$objWriter->endElement();
				}

				//	Labels
				$plotSeriesLabel = $plotGroup->getPlotLabelByIndex($plotSeriesRef);
				if ($plotSeriesLabel && ($plotSeriesLabel->getPointCount() > 0)) {
					$objWriter->startElement('c:tx');
						$objWriter->startElement('c:strRef');
							$this->_writePlotSeriesLabel($plotSeriesLabel, $objWriter);
						$objWriter->endElement();
					$objWriter->endElement();
				}
				
				// Color / Gradient Scheme
				if (isset($options['colors']) || isset($options['gradients'])) {
					if ($groupType !== PHPExcel_Chart_DataSeries::TYPE_LINECHART && $groupType !== PHPExcel_Chart_DataSeries::TYPE_STOCKCHART) {
						$objWriter->startElement('c:spPr');
							if (isset($options['gradients'])) {
								
								$gradient = $options['gradients'][$colorIx];
								$objWriter->startElement('a:gradFill');
									$objWriter->startElement('a:gsLst');
									foreach($gradient['fill'] as $gradientFill) {
										$objWriter->startElement('a:gs');
											$objWriter->writeAttribute('pos', $gradientFill['pos']);
											$objWriter->startElement('a:srgbClr');
												$objWriter->writeAttribute('val', strtoupper($gradientFill['srgb']));
												if (isset($gradientFill['alpha'])) {
													$objWriter->startElement('a:alpha');
														$objWriter->writeAttribute('val', strtoupper($gradientFill['alpha']));
													$objWriter->endElement();
												}
											$objWriter->endElement();
										$objWriter->endElement();
									}
									$objWriter->endElement();
									$objWriter->startElement('a:lin');
										$objWriter->writeAttribute('ang', $gradient['ang']);
										$objWriter->writeAttribute('scaled', 1);
									$objWriter->endElement();
								$objWriter->endElement();
								if (isset($gradient['line'])) {
									$objWriter->startElement('a:ln');
										$objWriter->writeAttribute('w', $gradient['line']['w']);
										$objWriter->startElement('a:solidFill');
											$objWriter->startElement('a:srgbClr');
												$objWriter->writeAttribute('val', strtoupper($gradient['line']['srgb']));
												if (isset($gradient['line']['alpha'])) {
													$objWriter->startElement('a:alpha');
														$objWriter->writeAttribute('val', strtoupper($gradient['line']['alpha']));
													$objWriter->endElement();
												}
											$objWriter->endElement();
										$objWriter->endElement();
									$objWriter->endElement();
								}
								
							} else if (isset($options['colors'])) {
								
								$color = $options['colors'][$colorIx];
								$objWriter->startElement('a:solidFill');
									$objWriter->startElement('a:srgbClr');
										$objWriter->writeAttribute('val', strtoupper($color['fill']['srgb']));
										if (isset($color['fill']['alpha'])) {
											$objWriter->startElement('a:alpha');
												$objWriter->writeAttribute('val', strtoupper($color['fill']['alpha']));
											$objWriter->endElement();
										}
									$objWriter->endElement();
								$objWriter->endElement();
								if (isset($color['line']['w'])) {
									$objWriter->startElement('a:ln');
										$objWriter->writeAttribute('w', $color['line']['w']);
										$objWriter->startElement('a:solidFill');
											$objWriter->startElement('a:srgbClr');
												$objWriter->writeAttribute('val', strtoupper($color['line']['srgb']));
												if (isset($color['line']['alpha'])) {
													$objWriter->startElement('a:alpha');
														$objWriter->writeAttribute('val', strtoupper($color['line']['alpha']));
													$objWriter->endElement();
												}
											$objWriter->endElement();
										$objWriter->endElement();
									$objWriter->endElement();
								}
								
							}
						$objWriter->endElement();
					}
				}

				//	Formatting for the points
				if (($groupType == PHPExcel_Chart_DataSeries::TYPE_LINECHART) ||
                    ($groupType == PHPExcel_Chart_DataSeries::TYPE_STOCKCHART)) {
					$objWriter->startElement('c:spPr');
						$objWriter->startElement('a:ln');
							$objWriter->writeAttribute('w', 12700);
            				if ($groupType == PHPExcel_Chart_DataSeries::TYPE_STOCKCHART) {
						        $objWriter->startElement('a:noFill');
						        $objWriter->endElement();
                            }

							if (isset($options['colors'])) {
								$objWriter->startElement('a:solidFill');
									$objWriter->startElement('a:srgbClr');
										$objWriter->writeAttribute('val', strtoupper($options['colors'][$colorIx]['fill']['srgb']));
									$objWriter->endElement();
								$objWriter->endElement();
							}

						$objWriter->endElement();
					$objWriter->endElement();
				}

				$plotSeriesValues = $plotGroup->getPlotValuesByIndex($plotSeriesRef);
				if ($plotSeriesValues) {
					$plotSeriesMarker = $plotSeriesValues->getPointMarker();
					if ($plotSeriesMarker) {
						$objWriter->startElement('c:marker');
							$objWriter->startElement('c:symbol');
								$objWriter->writeAttribute('val', $plotSeriesMarker);
							$objWriter->endElement();

							if ($plotSeriesMarker !== 'none') {
								$objWriter->startElement('c:size');
									$objWriter->writeAttribute('val', 3);
								$objWriter->endElement();
							}
						$objWriter->endElement();
					}
				}

				if (($groupType === PHPExcel_Chart_DataSeries::TYPE_BARCHART) ||
					($groupType === PHPExcel_Chart_DataSeries::TYPE_BARCHART_3D) ||
					($groupType === PHPExcel_Chart_DataSeries::TYPE_BUBBLECHART)) {

					$objWriter->startElement('c:invertIfNegative');
						$objWriter->writeAttribute('val', 0);
					$objWriter->endElement();
				}

				//	Category Labels
				$plotSeriesCategory = $plotGroup->getPlotCategoryByIndex($plotSeriesRef);
				if ($plotSeriesCategory && ($plotSeriesCategory->getPointCount() > 0)) {
					$catIsMultiLevelSeries = $catIsMultiLevelSeries || $plotSeriesCategory->isMultiLevelSeries();

					if (($groupType == PHPExcel_Chart_DataSeries::TYPE_PIECHART) ||
						($groupType == PHPExcel_Chart_DataSeries::TYPE_PIECHART_3D) ||
						($groupType == PHPExcel_Chart_DataSeries::TYPE_DONUTCHART)) {

						if (!is_null($plotGroup->getPlotStyle())) {
							$plotStyle = $plotGroup->getPlotStyle();
							if ($plotStyle) {
								$objWriter->startElement('c:explosion');
									$objWriter->writeAttribute('val', 25);
								$objWriter->endElement();
							}
						}
					}

					if (($groupType === PHPExcel_Chart_DataSeries::TYPE_BUBBLECHART) ||
						($groupType === PHPExcel_Chart_DataSeries::TYPE_SCATTERCHART)) {
						$objWriter->startElement('c:xVal');
					} else {
						$objWriter->startElement('c:cat');
					}

						$this->_writePlotSeriesValues($plotSeriesCategory, $objWriter, $groupType, 'str', $pSheet);
					$objWriter->endElement();
				}

				//	Values
				if ($plotSeriesValues) {
					$valIsMultiLevelSeries = $valIsMultiLevelSeries || $plotSeriesValues->isMultiLevelSeries();

					if (($groupType === PHPExcel_Chart_DataSeries::TYPE_BUBBLECHART) ||
						($groupType === PHPExcel_Chart_DataSeries::TYPE_SCATTERCHART)) {
						$objWriter->startElement('c:yVal');
					} else {
						$objWriter->startElement('c:val');
					}

						$this->_writePlotSeriesValues($plotSeriesValues, $objWriter, $groupType, 'num', $pSheet);
					$objWriter->endElement();
				}

				if ($groupType === PHPExcel_Chart_DataSeries::TYPE_BUBBLECHART) {
					$this->_writeBubbles($plotSeriesValues, $objWriter, $pSheet);
				}
				
				
				// Trendline
				if (isset($options['showTrendline']) && $options['showTrendline'] && (!isset($options['showTrendlineFor']) || in_array($plotSeriesIdx, $options['showTrendlineFor']))) {
					$colorIx++;
					$objWriter->startElement('c:trendline');
						$objWriter->startElement('c:name');
							$objWriter->writeRawData($options['trendline']['name']);
						$objWriter->endElement();
						$objWriter->startElement('c:trendlineType');
							$objWriter->writeAttribute('val', $options['trendline']['trendlineType']);
						$objWriter->endElement();
						if (isset($options['trendline']['intercept'])) {
							$objWriter->startElement('c:intercept');
								$objWriter->writeAttribute('val', $options['trendline']['intercept']);
							$objWriter->endElement();
						}
						$objWriter->startElement('c:dispRSqr');
							$objWriter->writeAttribute('val', $options['trendline']['dispRSqr']);
						$objWriter->endElement();
						$objWriter->startElement('c:dispEq');
							$objWriter->writeAttribute('val', $options['trendline']['dispEq']);
						$objWriter->endElement();
						$objWriter->startElement('c:spPr');
							if (isset($options['gradients'])) {
								
								$gradient = $options['gradients'][$colorIx];
								if (isset($gradient['line'])) {
									$objWriter->startElement('a:ln');
										$objWriter->writeAttribute('w', 20000); //$color['line']['w']);
										$objWriter->startElement('a:solidFill');
											$objWriter->startElement('a:srgbClr');
												$objWriter->writeAttribute('val', strtoupper($gradient['line']['srgb']));
												if (isset($gradient['line']['alpha'])) {
													$objWriter->startElement('a:alpha');
														$objWriter->writeAttribute('val', strtoupper($gradient['line']['alpha']));
													$objWriter->endElement();
												}
											$objWriter->endElement();
										$objWriter->endElement();
										$objWriter->startElement('a:prstDash');
											$objWriter->writeAttribute('val', 'sysDash');
										$objWriter->endElement();
									$objWriter->endElement();
								}
								
							} else if (isset($options['colors'])) {
								
								$color = $options['colors'][$colorIx];
								if (isset($color['line']['w'])) {
									$objWriter->startElement('a:ln');
										$objWriter->writeAttribute('w', 20000); //$color['line']['w']);
										$objWriter->startElement('a:solidFill');
											$objWriter->startElement('a:srgbClr');
												$objWriter->writeAttribute('val', strtoupper($color['line']['srgb']));
												if (isset($color['line']['alpha'])) {
													$objWriter->startElement('a:alpha');
														$objWriter->writeAttribute('val', strtoupper($color['line']['alpha']));
													$objWriter->endElement();
												}
											$objWriter->endElement();
										$objWriter->endElement();
										$objWriter->startElement('a:prstDash');
											$objWriter->writeAttribute('val', 'sysDash');
										$objWriter->endElement();
									$objWriter->endElement();
								}
								
							}
						$objWriter->endElement();
					$objWriter->endElement();
				}

			$objWriter->endElement();

			$colorIx++;
			
		}

		$this->_seriesIndex += $plotSeriesIdx + 1;
	}

	/**
	 * Write Plot Series Label
	 *
	 * @param	PHPExcel_Chart_DataSeriesValues		$plotSeriesLabel
	 * @param 	PHPExcel_Shared_XMLWriter 			$objWriter 			XML Writer
	 * @throws 	PHPExcel_Writer_Exception
	 */
	private function _writePlotSeriesLabel($plotSeriesLabel, $objWriter)
	{
		if (is_null($plotSeriesLabel)) {
			return;
		}

		$objWriter->startElement('c:f');
			$objWriter->writeRawData($plotSeriesLabel->getDataSource());
		$objWriter->endElement();

		$objWriter->startElement('c:strCache');
			$objWriter->startElement('c:ptCount');
				$objWriter->writeAttribute('val', $plotSeriesLabel->getPointCount() );
			$objWriter->endElement();

			foreach($plotSeriesLabel->getDataValues() as $plotLabelKey => $plotLabelValue) {
				$objWriter->startElement('c:pt');
					$objWriter->writeAttribute('idx', $plotLabelKey );

					$objWriter->startElement('c:v');
						$objWriter->writeRawData( $plotLabelValue );
					$objWriter->endElement();
				$objWriter->endElement();
			}
		$objWriter->endElement();

	}

	/**
	 * Write Plot Series Values
	 *
	 * @param	PHPExcel_Chart_DataSeriesValues		$plotSeriesValues
	 * @param 	PHPExcel_Shared_XMLWriter 			$objWriter 			XML Writer
	 * @param	string								$groupType			Type of plot for dataseries
	 * @param	string								$dataType			Datatype of series values
	 * @param	PHPExcel_Worksheet 					$pSheet
	 * @throws 	PHPExcel_Writer_Exception
	 */
	private function _writePlotSeriesValues( $plotSeriesValues,
											 $objWriter,
											 $groupType,
											 $dataType='str',
											 PHPExcel_Worksheet $pSheet
										   )
	{
		if (is_null($plotSeriesValues)) {
			return;
		}

		if ($plotSeriesValues->isMultiLevelSeries()) {
			$levelCount = $plotSeriesValues->multiLevelCount();

			$objWriter->startElement('c:multiLvlStrRef');

				$objWriter->startElement('c:f');
					$objWriter->writeRawData( $plotSeriesValues->getDataSource() );
				$objWriter->endElement();

				$objWriter->startElement('c:multiLvlStrCache');

					$objWriter->startElement('c:ptCount');
						$objWriter->writeAttribute('val', $plotSeriesValues->getPointCount() );
					$objWriter->endElement();

					for ($level = 0; $level < $levelCount; ++$level) {
						$objWriter->startElement('c:lvl');

						foreach($plotSeriesValues->getDataValues() as $plotSeriesKey => $plotSeriesValue) {
							if (isset($plotSeriesValue[$level])) {
								$objWriter->startElement('c:pt');
									$objWriter->writeAttribute('idx', $plotSeriesKey );

									$objWriter->startElement('c:v');
										$objWriter->writeRawData( $plotSeriesValue[$level] );
									$objWriter->endElement();
								$objWriter->endElement();
							}
						}

						$objWriter->endElement();
					}

				$objWriter->endElement();

			$objWriter->endElement();
		} else {
			$objWriter->startElement('c:'.$dataType.'Ref');

				$objWriter->startElement('c:f');
					$objWriter->writeRawData( $plotSeriesValues->getDataSource() );
				$objWriter->endElement();

				$objWriter->startElement('c:'.$dataType.'Cache');

					if (($groupType != PHPExcel_Chart_DataSeries::TYPE_PIECHART) &&
						($groupType != PHPExcel_Chart_DataSeries::TYPE_PIECHART_3D) &&
						($groupType != PHPExcel_Chart_DataSeries::TYPE_DONUTCHART)) {

						if (($plotSeriesValues->getFormatCode() !== NULL) &&
							($plotSeriesValues->getFormatCode() !== '')) {
							$objWriter->startElement('c:formatCode');
								$objWriter->writeRawData( $plotSeriesValues->getFormatCode() );
							$objWriter->endElement();
						}
					}

					$objWriter->startElement('c:ptCount');
						$objWriter->writeAttribute('val', $plotSeriesValues->getPointCount() );
					$objWriter->endElement();

					$dataValues = $plotSeriesValues->getDataValues();
					if (!empty($dataValues)) {
						if (is_array($dataValues)) {
							foreach($dataValues as $plotSeriesKey => $plotSeriesValue) {
								$objWriter->startElement('c:pt');
									$objWriter->writeAttribute('idx', $plotSeriesKey );

									$objWriter->startElement('c:v');
										$objWriter->writeRawData( $plotSeriesValue );
									$objWriter->endElement();
								$objWriter->endElement();
							}
						}
					}

				$objWriter->endElement();

			$objWriter->endElement();
		}
	}

	/**
	 * Write Bubble Chart Details
	 *
	 * @param	PHPExcel_Chart_DataSeriesValues		$plotSeriesValues
	 * @param 	PHPExcel_Shared_XMLWriter 			$objWriter 			XML Writer
	 * @throws 	PHPExcel_Writer_Exception
	 */
	private function _writeBubbles($plotSeriesValues, $objWriter, PHPExcel_Worksheet $pSheet)
	{
		if (is_null($plotSeriesValues)) {
			return;
		}

		$objWriter->startElement('c:bubbleSize');
			$objWriter->startElement('c:numLit');

				$objWriter->startElement('c:formatCode');
					$objWriter->writeRawData( 'General' );
				$objWriter->endElement();

				$objWriter->startElement('c:ptCount');
					$objWriter->writeAttribute('val', $plotSeriesValues->getPointCount() );
				$objWriter->endElement();

				$dataValues = $plotSeriesValues->getDataValues();
				if (!empty($dataValues)) {
					if (is_array($dataValues)) {
						foreach($dataValues as $plotSeriesKey => $plotSeriesValue) {
							$objWriter->startElement('c:pt');
								$objWriter->writeAttribute('idx', $plotSeriesKey );
								$objWriter->startElement('c:v');
									$objWriter->writeRawData( 1 );
								$objWriter->endElement();
							$objWriter->endElement();
						}
					}
				}

			$objWriter->endElement();
		$objWriter->endElement();

		$objWriter->startElement('c:bubble3D');
			$objWriter->writeAttribute('val', 0 );
		$objWriter->endElement();
	}

	/**
	 * Write Layout
	 *
	 * @param	PHPExcel_Chart_Layout		$layout
	 * @param 	PHPExcel_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @throws 	PHPExcel_Writer_Exception
	 */
	private function _writeLayout(PHPExcel_Chart_Layout $layout = NULL, $objWriter)
	{
		$objWriter->startElement('c:layout');

			if (!is_null($layout)) {
				$objWriter->startElement('c:manualLayout');

					$layoutTarget = $layout->getLayoutTarget();
					if (!is_null($layoutTarget)) {
						$objWriter->startElement('c:layoutTarget');
							$objWriter->writeAttribute('val', $layoutTarget);
						$objWriter->endElement();
					}

					$xMode = $layout->getXMode();
					if (!is_null($xMode)) {
						$objWriter->startElement('c:xMode');
							$objWriter->writeAttribute('val', $xMode);
						$objWriter->endElement();
					}

					$yMode = $layout->getYMode();
					if (!is_null($yMode)) {
						$objWriter->startElement('c:yMode');
							$objWriter->writeAttribute('val', $yMode);
						$objWriter->endElement();
					}

					$x = $layout->getXPosition();
					if (!is_null($x)) {
						$objWriter->startElement('c:x');
							$objWriter->writeAttribute('val', $x);
						$objWriter->endElement();
					}

					$y = $layout->getYPosition();
					if (!is_null($y)) {
						$objWriter->startElement('c:y');
							$objWriter->writeAttribute('val', $y);
						$objWriter->endElement();
					}

					$w = $layout->getWidth();
					if (!is_null($w)) {
						$objWriter->startElement('c:w');
							$objWriter->writeAttribute('val', $w);
						$objWriter->endElement();
					}

					$h = $layout->getHeight();
					if (!is_null($h)) {
						$objWriter->startElement('c:h');
							$objWriter->writeAttribute('val', $h);
						$objWriter->endElement();
					}

				$objWriter->endElement();
			}

		$objWriter->endElement();
	}

	/**
	 * Write Alternate Content block
	 *
	 * @param 	PHPExcel_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @throws 	PHPExcel_Writer_Exception
	 */
	private function _writeAlternateContent($objWriter)
	{
		$objWriter->startElement('mc:AlternateContent');
			$objWriter->writeAttribute('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006');

			$objWriter->startElement('mc:Choice');
				$objWriter->writeAttribute('xmlns:c14', 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart');
				$objWriter->writeAttribute('Requires', 'c14');

				$objWriter->startElement('c14:style');
					$objWriter->writeAttribute('val', '102');
				$objWriter->endElement();
			$objWriter->endElement();

			$objWriter->startElement('mc:Fallback');
				$objWriter->startElement('c:style');
					$objWriter->writeAttribute('val', '2');
				$objWriter->endElement();
			$objWriter->endElement();

		$objWriter->endElement();
	}

	/**
	 * Write Printer Settings
	 *
	 * @param 	PHPExcel_Shared_XMLWriter 	$objWriter 		XML Writer
	 * @throws 	PHPExcel_Writer_Exception
	 */
	private function _writePrintSettings($objWriter)
	{
		$objWriter->startElement('c:printSettings');

			$objWriter->startElement('c:headerFooter');
			$objWriter->endElement();

			$objWriter->startElement('c:pageMargins');
				$objWriter->writeAttribute('footer', 0.3);
				$objWriter->writeAttribute('header', 0.3);
				$objWriter->writeAttribute('r', 0.7);
				$objWriter->writeAttribute('l', 0.7);
				$objWriter->writeAttribute('t', 0.75);
				$objWriter->writeAttribute('b', 0.75);
			$objWriter->endElement();

			$objWriter->startElement('c:pageSetup');
				$objWriter->writeAttribute('orientation', "portrait");
			$objWriter->endElement();

		$objWriter->endElement();
	}

}
