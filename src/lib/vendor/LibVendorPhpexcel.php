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


/**
 * @package Webfrap
 * @subpackage ModBase
 * @author Dominik Bonsch <dominik.bonsch@webfrap.de>
 * @copyright Webfrap Developer Network <contact@webfrap.de>
 */
class LibVendorPhpexcel
  extends PHPExcel
{

	/**
	 * Create a new PHPExcel with one Worksheet
	 * @param string $title
	 * @param string $sheetClass
	 */
	public function __construct( $title = null, $sheetClass = null )
	{
		// Initialise worksheet collection and add one worksheet
		$this->_workSheetCollection = array();
		
		if( $sheetClass )
		{  
		  $this->_workSheetCollection[0] = new $sheetClass( $this, $title );
		}
		else
		{ 
		  $this->_workSheetCollection[0] = new PHPExcel_Worksheet( $this, $title );
		}
		  
		$this->_activeSheetIndex = 0;

		// Create document properties
		$this->_properties = new PHPExcel_DocumentProperties();

		// Create document security
		$this->_security = new PHPExcel_DocumentSecurity();

		// Set named ranges
		$this->_namedRanges = array();

		// Create the cellXf supervisor
		$this->_cellXfSupervisor = new PHPExcel_Style(true);
		$this->_cellXfSupervisor->bindParent($this);

		// Create the default style
		$this->addCellXf(new PHPExcel_Style);
		$this->addCellStyleXf(new PHPExcel_Style);
		
		
	}//end public function __construct */
  

}//end class LibVendorPhpexcel




