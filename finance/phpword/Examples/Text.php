<?php
require_once 'phpword/PHPWord.php';

// New Word Document
$PHPWord = new PHPWord();

// New portrait section
$section = $PHPWord->createSection();

// Add text elements
$section->addText('Cur');
$section->addTextBreak(2);

$section->addText('Name : Thomas ', array('name'=>'Verdana', 'color'=>'006699'));
$section->addTextBreak(2);

$PHPWord->addFontStyle('rStyle', array('bold'=>true, 'italic'=>true, 'size'=>16));
$PHPWord->addParagraphStyle('pStyle', array('align'=>'center', 'spaceAfter'=>100));
$section->addText('Address : 1 New york City', 'rStyle', 'pStyle');
$section->addText('', null, 'pStyle');

$section->addImage('photo1.jpg');
$section->addTextBreak(2);

// Save File
$objWriter = PHPWord_IOFactory::createWriter($PHPWord, 'Word2007');
$objWriter->save('Text.docx');
?>