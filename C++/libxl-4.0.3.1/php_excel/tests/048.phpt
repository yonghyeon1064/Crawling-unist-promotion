--TEST--
Column Read
--INI--
date.timezone=America/Toronto
--SKIPIF--
<?php if (!extension_loaded("excel")) print "skip"; ?>
--FILE--
<?php 
	$x = new ExcelBook();

	$s = $x->addSheet("Sheet 1");

	$data = array(true, 1.222, 434324, "fsdfasDF", NULL, "", false, -3321, -77.3321, "a a a a a aa");

	for ($i = 0; $i < 10; $i++) {
		for ($j = 0; $j < 10; $j++) {
			$s->write($j, $i, $data[$j]);
		}
	}

	var_dump($s->readCol(2), $x->getError());
	var_dump($s->readCol(2, 4), $x->getError());	
	var_dump($s->readCol(2, 5, 5), $x->getError());

	var_dump($s->readCol(-2));
	var_dump($s->readCol(22));
	var_dump($s->readCol(2, -1));
	var_dump($s->readCol(2, 55));
	var_dump($s->readCol(2, 2, 1));
	var_dump($s->readCol(2, 2, 39));
	
	echo "OK\n";
?>
--EXPECTF--
array(10) {
  [0]=>
  bool(true)
  [1]=>
  float(1.222)
  [2]=>
  float(434324)
  [3]=>
  string(8) "fsdfasDF"
  [4]=>
  NULL
  [5]=>
  string(0) ""
  [6]=>
  bool(false)
  [7]=>
  float(-3321)
  [8]=>
  float(-77.3321)
  [9]=>
  string(12) "a a a a a aa"
}
bool(false)
array(6) {
  [0]=>
  NULL
  [1]=>
  string(0) ""
  [2]=>
  bool(false)
  [3]=>
  float(-3321)
  [4]=>
  float(-77.3321)
  [5]=>
  string(12) "a a a a a aa"
}
bool(false)
array(1) {
  [0]=>
  string(0) ""
}
bool(false)

Warning: ExcelSheet::readCol(): Invalid column number '-2' in %s on line %d
bool(false)

Warning: ExcelSheet::readCol(): Invalid column number '22' in %s on line %d
bool(false)

Warning: ExcelSheet::readCol(): Invalid starting row number '-1' in %s on line %d
bool(false)

Warning: ExcelSheet::readCol(): Invalid starting row number '55' in %s on line %d
bool(false)

Warning: ExcelSheet::readCol(): Invalid ending row number '1' in %s on line %d
bool(false)

Warning: ExcelSheet::readCol(): Invalid ending row number '39' in %s on line %d
bool(false)
OK
