<?php
require_once 'phpword/PHPWord.php';
// New Word Document
$PHPWord = new PHPWord();
// New portrait section
$section = $PHPWord->createSection();
// Add image elements
$section->addImage('phpword/Examples/_mars.jpg');
$section->addTextBreak(2);
$section->addImage('phpword/Examples/_earth.JPG', array('width'=>210, 'height'=>210, 'align'=>'center'));
$section->addTextBreak(2);
$section->addImage('phpword/Examples/_mars.jpg', array('width'=>100, 'height'=>100, 'align'=>'right'));
// Save File
$objWriter = PHPWord_IOFactory::createWriter($PHPWord, 'Word2007');
$objWriter->save('Image.docx');
?>