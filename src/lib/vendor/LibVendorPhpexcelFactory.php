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

// special code block

if( !defined( 'VENDOR_PHPEXCEL_VERSION' ) )
{

  define( 'VENDOR_PHPEXCEL_VERSION' , 'stable' );

  // include the needed data
  include PATH_ROOT.'WebFrap_Lib_PHPExcel/vendor/phpexcel/'.VENDOR_PHPEXCEL_VERSION.'/Classes/PHPExcel.php';
  
  /** PHPExcel_Cell_AdvancedValueBinder */
  include PATH_ROOT.'WebFrap_Lib_PHPExcel/vendor/phpexcel/'.VENDOR_PHPEXCEL_VERSION.'/Classes/PHPExcel/Cell/AdvancedValueBinder.php';
  
  /** PHPExcel_IOFactory */
  include PATH_ROOT.'WebFrap_Lib_PHPExcel/vendor/phpexcel/'.VENDOR_PHPEXCEL_VERSION.'/Classes/PHPExcel/IOFactory.php';
  
  /** Read Filters */
  include PATH_ROOT.'WebFrap_Lib_PHPExcel/vendor/phpexcel/'.VENDOR_PHPEXCEL_VERSION.'/Classes/PHPExcel/Reader/IReadFilter.php';

}


// Set value binder
PHPExcel_Cell::setValueBinder( new PHPExcel_Cell_AdvancedValueBinder() );

/**
 * @package Webfrap
 * @subpackage ModBase
 * @author Dominik Bonsch <dominik.bonsch@webfrap.de>
 * @copyright Webfrap Developer Network <contact@webfrap.de>
 * @licence GPLv3
 */
class LibVendorPhpexcelFactory
{

  /**
   *
   * @var LibVendorPhpexcelFactory
   */
  private static $instance = null;

  /**
   * simple method just to call the autoload
   * @return LibVendorPhpexcelFactory
   */
  public static function getDefault()
  {

    if(!self::$instance)
      self::$instance = new LibVendorPhpexcelFactory();

    return self::$instance;

  }//end public static function getDefault */

/*//////////////////////////////////////////////////////////////////////////////
// Readers
//////////////////////////////////////////////////////////////////////////////*/
  
  /**
   * @param string $title
   * @param string $sheetClass
   */
  public static function newDocument( $title = null, $sheetClass = null )
  {
    
    return new LibVendorPhpexcelDocument( $title, $sheetClass );
    
  }//end public static function newDocument  */

/*//////////////////////////////////////////////////////////////////////////////
// Readers
//////////////////////////////////////////////////////////////////////////////*/

  /**
   *
   */
  public function loadReader2007()
  {

    // include the needed data
    include PATH_ROOT.'WebFrap_Lib_PHPExcel/vendor/phpexcel/'.VENDOR_PHPEXCEL_VERSION.'/Classes/PHPExcel/Reader/Excel2007.php';

  }//end public static function loadReader2007 */

  /**
   *
   */
  public function loadReader2003()
  {

    // include the needed data
    include PATH_ROOT.'WebFrap_Lib_PHPExcel/vendor/phpexcel/'.VENDOR_PHPEXCEL_VERSION.'/Classes/PHPExcel/Reader/Excel2003XML.php';

  }//end public function loadReader2003 */

  /**
   *
   */
  public function loadReader5()
  {

    // include the needed data
    include PATH_ROOT.'WebFrap_Lib_PHPExcel/vendor/phpexcel/'.VENDOR_PHPEXCEL_VERSION.'/Classes/PHPExcel/Reader/Excel5.php';

  }//end public function loadReader5 */

  /**
   *
   */
  public function loadReaderOOCalc()
  {

    // include the needed data
    include PATH_ROOT.'WebFrap_Lib_PHPExcel/vendor/phpexcel/'.VENDOR_PHPEXCEL_VERSION.'/Classes/PHPExcel/Reader/OOCalc.php';

  }//end public function loadReaderOOCalc */

/*//////////////////////////////////////////////////////////////////////////////
// Writers
//////////////////////////////////////////////////////////////////////////////*/

  /**
   *
   */
  public function loadWriter2007()
  {

    // include the needed data
    include PATH_ROOT.'WebFrap_Lib_PHPExcel/vendor/phpexcel/'.VENDOR_PHPEXCEL_VERSION.'/Classes/PHPExcel/Writer/Excel2007.php';

  }//end public function loadWriter2007 */

/*//////////////////////////////////////////////////////////////////////////////
// Some Loaders
//////////////////////////////////////////////////////////////////////////////*/

  /**
   *
   * @param PHPExcel $phpExcel
   */
  public function getExcelWriter2007( $phpExcel )
  {

    if( !class_exists('PHPExcel_Writer_Excel2007') )
      self::loadWriter2007();

    return PHPExcel_IOFactory::createWriter($phpExcel, "Excel2007");

  }//end public function getExcelWriter2007 */

  /**
   * @param string $title
   * @param string $sheetClass
   */
  public function newSheet( $title = null, $sheetClass = null )
  {
    
    return new LibVendorPhpexcelDocument( $title, $sheetClass );
    
  }//end public function newSheet  */


}//end class LibVendorPhpexcelFactory




