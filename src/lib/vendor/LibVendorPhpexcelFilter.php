<?php
/*******************************************************************************
*
* @author      : Dominik Bonsch <dominik.bonsch@webfrap.net>
* @date        :
* @copyright   : Webfrap Developer Network <contact@webfrap.net>
* @project     : Webfrap Web Frame Application
* @projectUrl  : http://webfrap.net
*
* @licence     : BSD License see: LICENCE/BSD Licence.txt
* 
* @version: @package_version@  Revision: @package_revision@
*
* Changes:
*
*******************************************************************************/


class LibVendorPhpexcelFilter
  implements PHPExcel_Reader_IReadFilter
{

  /**
   *
   */
  public $start     = 1;

  /**
   *
   */
  public $end       = 2500;

  /**
   *
   */
  public $step      = 2500;

  /**
   *
   */
  public $changed   = false;

  /**
   *
   */
  public function setStart( $start )
  {
    $this->start  = $start;
    $this->end    = $start + $this->step;
  }

  /**
   *
   */
  public function readCell($column, $row, $worksheetName = '')
  {

    // Read title row and rows 20 - 30
    if ( ($row > $this->start && $row <= $this->end ))
    {
      $this->changed = true;
      return true;
    }

    return false;

  }//end public function readCell */

}//end class LibVendorPhpexcelFilter */
