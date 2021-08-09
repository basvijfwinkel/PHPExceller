<?php
namespace PHPExceller\Chart;

use PHPExceller\Chart\ChartProperties;

/**
 * Based on PHPExcel_Chart_Axis
 */

class Axis extends  ChartProperties 
{

  /**
   * Axis Number
   *
   * @var  array of mixed
   */

  private
      $axis_number = array(
      'format' => self::FORMAT_CODE_GENERAL,
      'source_linked' => 1
  );

  /**
   *
   *
   *
   * @var rotation of the labels of the axis in degrees
   */
   private $axis_text_rotation = 0;

  /**
   * Axis Options
   *
   * @var  array of mixed
   */

  private $axis_options = array(
      'minimum' => NULL,
      'maximum' => NULL,
      'major_unit' => NULL,
      'minor_unit' => NULL,
      'orientation' => self::ORIENTATION_NORMAL,
      'minor_tick_mark' => self::TICK_MARK_NONE,
      'major_tick_mark' => self::TICK_MARK_NONE,
      'axis_labels' => self::AXIS_LABELS_NEXT_TO,
      'horizontal_crosses' => self::HORIZONTAL_CROSSES_AUTOZERO,
      'horizontal_crosses_value' => NULL,
      'axis_position' => 'l',
      'hide_major_gridlines' => false,
  );

  /**
   * Fill Properties
   *
   * @var  array of mixed
   */

  private $fill_properties = array(
      'type' => self::EXCEL_COLOR_TYPE_ARGB,
      'value' => NULL,
      'alpha' => 0
  );

  /**
   * Line Properties
   *
   * @var  array of mixed
   */

  private $line_properties = array(
      'type' => self::EXCEL_COLOR_TYPE_ARGB,
      'value' => NULL,
      'alpha' => 0
  );

  /**
   * Line Style Properties
   *
   * @var  array of mixed
   */

  private $line_style_properties = array(
      'width' => '9525',
      'compound' => self::LINE_STYLE_COMPOUND_SIMPLE,
      'dash' => self::LINE_STYLE_DASH_SOLID,
      'cap' => self::LINE_STYLE_CAP_FLAT,
      'join' => self::LINE_STYLE_JOIN_BEVEL,
      'arrow' => array(
          'head' => array(
              'type' => self::LINE_STYLE_ARROW_TYPE_NOARROW,
              'size' => self::LINE_STYLE_ARROW_SIZE_5
          ),
          'end' => array(
              'type' => self::LINE_STYLE_ARROW_TYPE_NOARROW,
              'size' => self::LINE_STYLE_ARROW_SIZE_8
          ),
      )
  );

  /**
   * Shadow Properties
   *
   * @var  array of mixed
   */

  private $shadow_properties = array(
      'presets' => self::SHADOW_PRESETS_NOSHADOW,
      'effect' => NULL,
      'color' => array(
          'type' => self::EXCEL_COLOR_TYPE_STANDARD,
          'value' => 'black',
          'alpha' => 40,
      ),
      'size' => array(
          'sx' => NULL,
          'sy' => NULL,
          'kx' => NULL
      ),
      'blur' => NULL,
      'direction' => NULL,
      'distance' => NULL,
      'algn' => NULL,
      'rotWithShape' => NULL
  );

  /**
   * Glow Properties
   *
   * @var  array of mixed
   */

  private $glow_properties = array(
      'size' => NULL,
      'color' => array(
          'type' => self::EXCEL_COLOR_TYPE_STANDARD,
          'value' => 'black',
          'alpha' => 40
      )
  );

  /**
   * Soft Edge Properties
   *
   * @var  array of mixed
   */

  private $soft_edges = array(
      'size' => NULL
  );

  /**
   * Get Series Data Type
   *
   * @return  string
   */

  public function setAxisNumberProperties($format_code) {
    $this->axis_number['format'] = (string) $format_code;
    $this->axis_number['source_linked'] = 0;
  }

  /**
   * Get Axis Number Format Data Type
   *
   * @return  string
   */

  public function getAxisNumberFormat() {
    return $this->axis_number['format'];
  }

  /**
   * Get Axis Number Source Linked
   *
   * @return  string
   */

  public function getAxisNumberSourceLinked() {
    return (string) $this->axis_number['source_linked'];
  }

  /**
   * Set Axis Options Properties
   *
   * @param string $axis_labels
   * @param string $horizontal_crosses_value
   * @param string $horizontal_crosses
   * @param string $axis_orientation
   * @param string $major_tmt
   * @param string $minor_tmt
   * @param string $minimum
   * @param string $maximum
   * @param string $major_unit
   * @param string $minor_unit
   *
   */

  public function setAxisOptionsProperties($axis_labels, $horizontal_crosses_value = NULL, $horizontal_crosses = NULL,
      $axis_orientation = NULL, $major_tmt = NULL, $minor_tmt = NULL, $minimum = NULL, $maximum = NULL, $major_unit = NULL,
      $minor_unit = NULL,$axis_position=NULL, $hide_major_gridlines=NULL) {

    $this->axis_options['axis_labels'] = (string) $axis_labels;
    ($horizontal_crosses_value !== NULL)
        ? $this->axis_options['horizontal_crosses_value'] = (string) $horizontal_crosses_value : NULL;
    ($horizontal_crosses !== NULL) ? $this->axis_options['horizontal_crosses'] = (string) $horizontal_crosses : NULL;
    ($axis_orientation !== NULL) ? $this->axis_options['orientation'] = (string) $axis_orientation : NULL;
    ($major_tmt !== NULL) ? $this->axis_options['major_tick_mark'] = (string) $major_tmt : NULL;
    ($minor_tmt !== NULL) ? $this->axis_options['minor_tick_mark'] = (string) $minor_tmt : NULL;
    ($minor_tmt !== NULL) ? $this->axis_options['minor_tick_mark'] = (string) $minor_tmt : NULL;
    ($minimum !== NULL) ? $this->axis_options['minimum'] = (string) $minimum : NULL;
    ($maximum !== NULL) ? $this->axis_options['maximum'] = (string) $maximum : NULL;
    ($major_unit !== NULL) ? $this->axis_options['major_unit'] = (string) $major_unit : NULL;
    ($minor_unit !== NULL) ? $this->axis_options['minor_unit'] = (string) $minor_unit : NULL;
    ($axis_position !== NULL) ? $this->axis_options['axis_position'] = (string) $axis_position : NULL;
    ($hide_major_gridlines !== NULL) ? $this->axis_options['hide_major_gridlines'] = $hide_major_gridlines : FALSE;
  }
  
  public function setAxisOptionsProperty($name, $value)
  {
      $this->axis_options[$name] = $value;
  }

  /**
   * Get Axis Options Property
   *
   * @param string $property
   *
   * @return string
   */

  public function getAxisOptionsProperty($property) {
    return $this->axis_options[$property];
  }

  /**
   * Set Axis Orientation Property
   *
   * @param string $orientation
   *
   */

  public function setAxisOrientation($orientation) {
    $this->orientation = (string) $orientation;
  }

  /**
   * Set Fill Property
   *
   * @param string $color
   * @param int $alpha
   * @param string $type
   *
   */

  public function setFillParameters($color, $alpha = 0, $type = self::EXCEL_COLOR_TYPE_ARGB) {
    $this->fill_properties = $this->setColorProperties($color, $alpha, $type);
  }

  /**
   * Set Line Property
   *
   * @param string $color
   * @param int $alpha
   * @param string $type
   *
   */

  public function setLineParameters($color, $alpha = 0, $type = self::EXCEL_COLOR_TYPE_ARGB) {
    $this->line_properties = $this->setColorProperties($color, $alpha, $type);
  }

  /**
   * Get Fill Property
   *
   * @param string $property
   *
   * @return string
   */

  public function getFillProperty($property) {
    return $this->fill_properties[$property];
  }

  /**
   * Get Line Property
   *
   * @param string $property
   *
   * @return string
   */

  public function getLineProperty($property) {
    return $this->line_properties[$property];
  }

  /**
   * Set Line Style Properties
   *
   * @param float $line_width
   * @param string $compound_type
   * @param string $dash_type
   * @param string $cap_type
   * @param string $join_type
   * @param string $head_arrow_type
   * @param string $head_arrow_size
   * @param string $end_arrow_type
   * @param string $end_arrow_size
   *
   */

  public function setLineStyleProperties($line_width = NULL, $compound_type = NULL,
      $dash_type = NULL, $cap_type = NULL, $join_type = NULL, $head_arrow_type = NULL,
      $head_arrow_size = NULL, $end_arrow_type = NULL, $end_arrow_size = NULL) {

    (!is_null($line_width)) ? $this->line_style_properties['width'] = $this->getExcelPointsWidth((float) $line_width)
        : NULL;
    (!is_null($compound_type)) ? $this->line_style_properties['compound'] = (string) $compound_type : NULL;
    (!is_null($dash_type)) ? $this->line_style_properties['dash'] = (string) $dash_type : NULL;
    (!is_null($cap_type)) ? $this->line_style_properties['cap'] = (string) $cap_type : NULL;
    (!is_null($join_type)) ? $this->line_style_properties['join'] = (string) $join_type : NULL;
    (!is_null($head_arrow_type)) ? $this->line_style_properties['arrow']['head']['type'] = (string) $head_arrow_type
        : NULL;
    (!is_null($head_arrow_size)) ? $this->line_style_properties['arrow']['head']['size'] = (string) $head_arrow_size
        : NULL;
    (!is_null($end_arrow_type)) ? $this->line_style_properties['arrow']['end']['type'] = (string) $end_arrow_type
        : NULL;
    (!is_null($end_arrow_size)) ? $this->line_style_properties['arrow']['end']['size'] = (string) $end_arrow_size
        : NULL;
  }

  /**
   * Get Line Style Property
   *
   * @param array|string $elements
   *
   * @return string
   */

  public function getLineStyleProperty($elements) {
    return $this->getArrayElementsValue($this->line_style_properties, $elements);
  }

  /**
   * Get Line Style Arrow Excel Width
   *
   * @param string $arrow
   *
   * @return string
   */

  public function getLineStyleArrowWidth($arrow) {
    return $this->getLineStyleArrowSize($this->line_style_properties['arrow'][$arrow]['size'], 'w');
  }

  /**
   * Get Line Style Arrow Excel Length
   *
   * @param string $arrow
   *
   * @return string
   */

  public function getLineStyleArrowLength($arrow) {
    return $this->getLineStyleArrowSize($this->line_style_properties['arrow'][$arrow]['size'], 'len');
  }

  /**
   * Set Shadow Properties
   *
   * @param int $shadow_presets
   * @param string $sh_color_value
   * @param string $sh_color_type
   * @param string $sh_color_alpha
   * @param float $sh_blur
   * @param int $sh_angle
   * @param float $sh_distance
   *
   */

  public function setShadowProperties($sh_presets, $sh_color_value = NULL, $sh_color_type = NULL, $sh_color_alpha = NULL, $sh_blur = NULL, $sh_angle = NULL, $sh_distance = NULL) {
    $this
        ->_setShadowPresetsProperties((int) $sh_presets)
        ->_setShadowColor(
            is_null($sh_color_value) ? $this->shadow_properties['color']['value'] : $sh_color_value
            , is_null($sh_color_alpha) ? (int) $this->shadow_properties['color']['alpha'] : $sh_color_alpha
            , is_null($sh_color_type) ? $this->shadow_properties['color']['type'] : $sh_color_type)
        ->_setShadowBlur($sh_blur)
        ->_setShadowAngle($sh_angle)
        ->_setShadowDistance($sh_distance);
  }

  /**
   * Set Shadow Color
   *
   * @param int $shadow_presets
   *
   * @return void
   */

  private function _setShadowPresetsProperties($shadow_presets) {
    $this->shadow_properties['presets'] = $shadow_presets;
    $this->setShadowProperiesMapValues($this->getShadowPresetsMap($shadow_presets));
  }

  /**
   * Set Shadow Properties from Maped Values
   *
   * @param array $properties_map
   * @param * $reference
   *
   * @return Axis
   */

  private function _setShadowProperiesMapValues(array $properties_map, &$reference = NULL) {
    $base_reference = $reference;
    foreach ($properties_map as $property_key => $property_val) {
      if (is_array($property_val)) {
        if ($reference === NULL) {
          $reference = & $this->shadow_properties[$property_key];
        } else {
          $reference = & $reference[$property_key];

    /**
     * Set Shadow Color
     *
     * @param string $color
     * @param int $alpha
     * @param string $type
     *
     * @return void
     */
    private function setShadowColor($color, $alpha, $type)
    {
        $this->shadow_properties['color'] = $this->setColorProperties($color, $alpha, $type);
    }

    /**
     * Set Shadow Blur
     *
     * @param float $blur
     *
     * @return void
     */
    private function setShadowBlur($blur)
    {
        if ($blur !== null) {
            $this->shadow_properties['blur'] = (string) $this->getExcelPointsWidth($blur);
        }
    }

    /**
     * Set Shadow Angle
     *
     * @param int $angle
     *
     * @return void
     */
    private function setShadowAngle($angle)
    {
        if ($angle !== null) {
            $this->shadow_roperties['direction'] = (string) $this->getExcelPointsAngle($angle);
        }
    }

    /**
     * Set Shadow Distance
     *
     * @param float $distance
     *
     * @return void
     */
    private function setShadowDistance($distance)
    {
        if ($distance !== null) {
            $this->shadow_properties['distance'] = (string) $this->getExcelPointsWidth($distance);
        }
    }

    /**
     * Get Glow Property
     *
     * @param float $size
     * @param string $color_value
     * @param int $color_alpha
     * @param string $color_type
     */
    public function getShadowProperty($elements)
    {
        return $this->getArrayElementsValue($this->shadow_properties, $elements);
    }

    /**
     * Set Glow Properties
     *
     * @param float $size
     * @param string $color_value
     * @param int $color_alpha
     * @param string $color_type
     */
    public function setGlowProperties($size, $color_value = null, $color_alpha = null, $color_type = null)
    {
        $this->setGlowSize($size)
            ->setGlowColor(
                is_null($color_value) ? $this->glow_properties['color']['value'] : $color_value,
                is_null($color_alpha) ? (int) $this->glow_properties['color']['alpha'] : $color_alpha,
                is_null($color_type) ? $this->glow_properties['color']['type'] : $color_type
            );
    }

    /**
     * Get Glow Property
     *
     * @param array|string $property
     *
     * @return string
     */
    public function getGlowProperty($property)
    {
        return $this->getArrayElementsValue($this->glow_properties, $property);
    }

    /**
     * Set Glow Color
     *
     * @param float $size
     *
     * @return void
     */
    private function setGlowSize($size)
    {
        if (!is_null($size)) {
            $this->glow_properties['size'] = $this->getExcelPointsWidth($size);
        }
    }

    /**
     * Set Glow Color
     *
     * @param string $color
     * @param int $alpha
     * @param string $type
     *
     * @return void
     */
    private function setGlowColor($color, $alpha, $type)
    {
        $this->glow_properties['color'] = $this->setColorProperties($color, $alpha, $type);
    }

    /**
     * Set Soft Edges Size
     *
     * @param float $size
     */
    public function setSoftEdges($size)
    {
        if (!is_null($size)) {
            $this->soft_edges['size'] = (string) $this->getExcelPointsWidth($size);
        }
    }

    /**
     * Get Soft Edges Size
     *
     * @return string
     */

    public function getSoftEdgesSize() {
       return $this->soft_edges['size'];
    }

    public function getAxisTextRotation() {
      return $this->axis_text_rotation;
    }

    public function setAxisTextRotation($rotation) {
      $this->axis_text_rotation = $rotation;
    }
}
