<?
IncludeModuleLangFile(__FILE__);

/**
 * Class ParserClass
 */
class ParserClass
{

	var $MODULE_ID = 'pixel.parser';

	/**
	 * @param $by
	 * @param $order
	 * @param $arFilter
	 *
	 * @return CDBResult
	 *
	 * пока не знаю для чего эта функция, но может пригодиться
	 */
	function GetList($by, $order, $arFilter)
	{
		global $DB;

		$query = $DB->Query('SELECT "NAME, CHANGED" FROM p_parser_res_list');

		while ($res = $query->Fetch()) {
			$elem['ID'] = $res['ID'];
			$elem['NAME'] = $res['NAME'];
			$elem['ACTIVE'] = $res['ACTIVE'];
			$elem['DATE_ACTIVE_FROM'] = $res['DATE_ACTIVE_FROM'];
			$elem['DATE_ACTIVE_TO'] = $res['DATE_ACTIVE_TO'];
			$arResult[] = $elem;
		}

		$result = new CDBResult;
		$result->InitFromArray($arResult);

		return $result;
	}//end function GetList

	/**
	 * @param $INPUT
	 *
	 * @return bool
	 *
	 * ф-я проверяет расширение файла на csv
	 */
	function validateFile($INPUT)
	{
		$ext = explode(".", $INPUT['filename']['name']);
		if($ext[1] != 'csv') {
			return false;
		} else return true;
	} //end function validateFile


	/**
	 * @param array $input
	 *
	 * @return bool
	 *
	 * функция загружает CSV
	 */
	function uploadCSV($input = Array())
	{

		if($input["filename"]["size"] > 1024 * 30 * 1024) {
			echo("Размер файла превышает три мегабайта");
			exit;
		}
		// Проверяем загружен ли файл
		if(is_uploaded_file($input["filename"]["tmp_name"])) {
			// Если файл загружен успешно, перемещаем его
			// из временной директории в конечную
			move_uploaded_file($input["filename"]["tmp_name"], $_SERVER['DOCUMENT_ROOT'] . "/upload/pixel.parser/" . $input["filename"]["name"]);
			chmod($_SERVER['DOCUMENT_ROOT'] . "/upload/pixel.parser/" . $input["filename"]["name"], 0775);
			chown($_SERVER['DOCUMENT_ROOT'] . "/upload/pixel.parser/" . $input["filename"]["name"], 'ned:root');
			return true;
		} else {
			return false;
		}
	}//end function uploadCSV

	/**
	 * @param $csv
	 *
	 * @return array
	 *
	 * функция читает CSV
	 */
	function readCSV($csv)
	{
		$arr = Array(); // array, into save values from uploaded csv
		$row = 1; // counter for csv rows
		$countRows = 1000; //ограничение на кол-во строк в CSV файле
		$file = fopen($csv, "r");// открываем файл

		// разбираем файл построчно
		while (($data = fgetcsv($file, $countRows, ";")) !== false) {
			//$countPlace = count($data); //quantity field in row (not used)
			$arr[] = $data;
			$row++;
		}
		fclose($file);
		return $arr;
	} //end function readCSV

	/**
	 * @param       $INPUT
	 * @param array $arr
	 * @param       $ITER
	 * @param       $error
	 *
	 * @return array|mixed
	 *
	 * парсинг http://domknigi-online.ru
	 */
	function goParseDomknigi($INPUT, $arr = Array(), $ITER, $error)
	{
		if((isset($INPUT)) && (!empty($INPUT)) && ($ITER >= 1) && (!isset($error))) {
			$rest = Array();
			foreach ($arr as $k => $v) {
				//name of search item
				$title = '';
				$nameOfBook = trim($v[1]);
				$stringOfSearch = str_replace(" ", "+", $nameOfBook);
				$stringOfSearch = mb_substr($stringOfSearch, 0, 20, utf8);
				//echo '$stringOfSearch: '.$stringOfSearch.'<br />';
				$result = $this->curl_get('http://domknigi-online.ru/index.php?searchword=' . $stringOfSearch . '&ordering=newest&searchphrase=all&option=com_search');
				$document = phpQuery::newDocument($result);
				$div_elem = $document->find(".contentpaneopen fieldset");
				$res = Array();
				foreach ($div_elem as $div) //обход блоков с анонсом в поисках нужного ISBN
				{
					//todo добаувить комменты что тут заносится в переменные phpQuery
					$pq = pq($div);
					$t = $pq->find("a");
					$cat = $pq->find("span.small:last");
					$cat = $cat->text();
					$title = $t->text();
					$href = $t->attr('href');

					// фильтр по категориям
					if(trim($cat) != '(Каталог/Новинки)' && trim($cat) != '(Каталог/Аудиокниги)' && trim($cat) != '(News/Latest)') {
						//todo плохо, сократить кол-во операций
						$title = str_replace(".", "", $title); // удаляем '.'
						$title = str_replace(",", "", $title); // удаляем ','
						$title = str_replace(":", "", $title); // удаляем ':'
						$title = str_replace(";", "", $title); // удаляем ';'
						$nameOfBook = str_replace(".", "", $nameOfBook); // удаляем '.'
						$nameOfBook = str_replace(",", "", $nameOfBook); // удаляем ','
						$nameOfBook = str_replace(":", "", $nameOfBook); // удаляем ':'
						$nameOfBook = str_replace(";", "", $nameOfBook); // удаляем ';'

						// проверка вхождения поиска в title и наоборот
						$match = stristr(trim($title), trim($nameOfBook));
						if($match != false) {
							// парсинг цены
							$result1 = $this->curl_get('http://domknigi-online.ru' . $href);
							$document1 = phpQuery::newDocument($result1);
							$div_elem1 = $document1->find(".productPrice");
							$price = $div_elem1->text();
							//todo сократить кол-во операций, проверить правильность логики работы
							$price = explode('00', $price);
							$price = preg_replace('/\D/', '', $price[0]);
							$res[] = $price;
							$res[] = $v[1];
							continue;

						} else {
							$match = stristr(trim($nameOfBook), trim($title));
							if($match != false) {
								// парсинг цены
								$result1 = $this->curl_get('http://domknigi-online.ru' . $href);
								$document1 = phpQuery::newDocument($result1);
								$div_elem1 = $document1->find(".productPrice");
								$price = $div_elem1->text();
								$price = explode('00', $price);
								$price = preg_replace('/\D/', '', $price[0]);
								$res[] = $price;
								$res[] = $v[1];
								continue;
							}
						}

					}
				} // end foreach
				sleep(0.1);
				if(!empty($res)) {
					$rest[] = $res;
				} else {
					$search_str = explode("+", $stringOfSearch);
					foreach ($search_str as $str) {
						if($title != '' || $str != '') {
							$match = stristr(trim($title), trim($str));
							if($match != false) {
								$result1 = $this->curl_get('http://domknigi-online.ru' . $href);
								$document1 = phpQuery::newDocument($result1);
								$div_elem1 = $document1->find(".productPrice");
								$price = $div_elem1->text();
								//todo сократить кол-вл операций
								$price = explode('00', $price);
								$price = preg_replace('/\D/', '', $price[0]);
								$res[] = $price;
								$res[] = $v[1];
								$rest[] = $res;
								break;
							}
						}
					}
				}

			}
			$result = array();
			foreach ($rest as $rst) {
				$result[] = array_unique($rst);
			}
			return $result;
		}

	} //end function goParseDomknigi

	/**
	 * @param       $INPUT
	 * @param array $arr
	 * @param       $ITER
	 * @param       $error
	 *
	 * @return array|bool
	 */
	function goParse_100000_knig($INPUT, $arr = Array(), $ITER, $error)
	{
		$flag = 0;
		if((isset($INPUT)) && (!empty($INPUT)) && ($ITER >= 1) && (!isset($error))) {
			$rest = Array();
			foreach ($arr as $k => $v) {
				$nameOfBook = trim(substr($v[1], 0, 90));
				if(strlen($nameOfBook) > 45) {
					// если более 45 симв., режем до пробела
					$nameOfBook = substr($nameOfBook, 0, strrpos($nameOfBook, ' '));
				}
				$len = explode(' ', $nameOfBook); // array of word name
				$stringOfSearch = '';
				$i = 0;
				while (($i <= (count($len) - 1)) && ($flag == 0)) {
					if($len[$i] != '') {
						$stringOfSearch .= urlencode(iconv('UTF8', 'CP1251', $len[$i])) . '+';
					}
					$stringOfSearchGo = substr($stringOfSearch, 0, strrpos($stringOfSearch, '+')); // режем посл. +
					$str_search = 'http://www.100000-knig.ru/catalog/search?author=%E0%E2%F2%EE%F0&name=' . $stringOfSearchGo . '&publisher=%E8%E7%E4%E0%F2%E5%EB%FC%F1%F2%E2%EE&Search=%C8%F1%EA%E0%F2%FC';
					$result = $this->curl_get($str_search);
					$document = phpQuery::newDocument($result);
					$div_elem = $document->find(".smallbook-zakaz");
					$res = Array();
					foreach ($div_elem as $div) //обход блоков с анонсом в поисках нужного ISBN
					{
						$pq = pq($div);
						$t = $pq->find("h3 > a");
						$href = $t->attr('href'); // ссылка на детальную стр.
						$isbn = $pq->find("ul > li:contains('ISBN:')");
						$isbn = $isbn->text();
						$isbn = explode('ISBN: ', $isbn);
						$isbn = preg_replace('/\D/', '', $isbn[1]); // чистим isbn

						// фильтр по ISBN-10 && ISBN-13
						if($isbn == $v[0] || $isbn == substr($v[0], 3, strlen($v[0]))) { //if ISBN equal means to find true
							$resultOfPrice = $this->curl_get('http://www.100000-knig.ru' . $href);
							$document1 = phpQuery::newDocument($resultOfPrice);
							$pr = $document1->find("div#tovar-price");
							$price = $pr->text();
							$price = explode(',00', $price);
							$price = preg_replace('/\D/', '', $price[0]);
							$res[] = $price;
							$res[] = $v[1];

							//echo $v[1].' - '.$price.'<br />';
							// todo записывать в CSV результаты
							//echo "<br>";
							//echo 'ISBN: '.$isbn .'<br />';
							$flag = 1;
							continue;
						}
					} // end foreach
					$i++;
					sleep(0.1);
				}// end while
				$rest[] = $res;
				$flag = 0;
			} //end foreach
			return $rest;
		} //end if
		else return false;
	}//end function goParse_100000_knig

	/**
	 * @param       $INPUT
	 * @param array $arr
	 * @param       $ITER
	 * @param       $error
	 *
	 * @return array|bool
	 *
	 * парсит mdk arbat
	 */
	function goParseMdkArbat($INPUT, $arr = Array(), $ITER, $error)
	{
		$flag = 0;

		if((isset($INPUT)) && (!empty($INPUT)) && ($ITER >= 1) && (!isset($error))) {
			$rest = Array();
			foreach ($arr as $k => $v) {
				$nameOfBook = trim(substr($v[1], 0, 70));
				$len = explode(' ', $nameOfBook);
				$len = count($len);
				while (($len > 1) && ($flag == 0)){

					$stringOfSearch = urlencode(iconv('UTF8', 'CP1251', $nameOfBook));

					$result = $this->curl_get('http://mdk-arbat.ru/catalog?kw=' . $stringOfSearch);
					$document = phpQuery::newDocument($result);
					$div_elem = $document->find(".good_description");

					$res = Array();
					foreach ($div_elem as $div) //обход блоков с анонсом в поисках нужного ISBN
					{
						$pq = pq($div);
						$t = $pq->find(".title");
						$tt = $t->text();

						$isb = $pq->find(".isbn");
						$isbn = pq($isb);
						$isbn = $isbn->text();
						$isbn = preg_replace('/\D/', '', $isbn);

						if($isbn == $v[0]) { //if ISBN equal means to find true
							$pr = $pq->find(".price_info");
							$price = pq($pr);
							$price = $price->text();
							$price = preg_replace('/\D/', '', $price);
							$res[] = $price;
							$res[] = $v[1];
							//echo $v[1].' - '.$price;
							// todo записывать в CSV результаты
							//echo "<br>";
							$flag = 1;
							continue;
						}
					} // end foreach

					sleep(0.1);
					//reduce the length of a search query in one word
					$nameOfBook = preg_replace('/\s\S+$/', "", $nameOfBook);
					//recalculate $len
					$len = count(explode(" ", $nameOfBook));
				}; // end while
				$rest[] = $res;
				$flag = 0;
			} //end foreach

			return $rest;
		} //end if
		else return false;
	}//end function goParseMdkArbat

    function goParse_respublica($INPUT, $arr = Array(), $ITER, $error)
    {
        $flag = 0;
        if((isset($INPUT)) && (!empty($INPUT)) && ($ITER >= 1) && (!isset($error)))
        {
        $rest = Array(); // массив для возврата
        foreach ($arr as $k => $v)
        {
            $nameOfBook = trim(substr($v[1], 0, 70));
            $len = explode(' ', $nameOfBook);
            $len = count($len);
            $stringOfSearch = '';

            while (($len > 1) && ($flag == 0))
            {
                $res = array();
                //echo '$nameOfBook: '.$nameOfBook.'<br />';
                $stringOfSearch .= urlencode($nameOfBook).'+';
                $stringOfSearchGo = substr($stringOfSearch, 0, strrpos($stringOfSearch, '+')); // режем посл. +
                //echo '$stringOfSearchGo: '.$stringOfSearchGo.'<br />';

                $result = $this->curl_get('http://www.respublica.ru/search?utf8=&filter[query]=' . $stringOfSearch);
                $document = phpQuery::newDocument($result);
                $div_elem = $document->find(".b-items-container.cfix");
                $price = $div_elem->find(".b-item-price-normal.i-left");
                if(count($price) == 1)
                {
                    $price = $price->text();
                    $price = explode('р', $price);
                    $price = preg_replace('/\D/', '', $price[0]);
                    $res[] = trim($price);
                    $res[] = $v[1];
                }
                else
                {
                    $res[] = 'нет';
                    $res[] = $v[1];
                }
                sleep(0.1);
                //reduce the length of a search query in one word
                $nameOfBook = preg_replace('/\s\S+$/', "", $nameOfBook);
                //recalculate $len
                $len = count(explode(" ", $nameOfBook));
                $stringOfSearch = '';
            }; // end while

            $rest[] = $res;
            $flag = 0;
        } //end foreach
        return $rest;
        } //end if
        else return false;
    }

	/**
	 * @param $host
	 *
	 * @return mixed
	 */
	function curl_get($host)
	{
		$ch = curl_init();
		if($_POST['PROXY'] != "") $proxy = $_POST['PROXY']; else $proxy = '112.25.12.37:80';
		//$proxyauth = 'user:password';
		curl_setopt($ch, CURLOPT_HEADER, 0);
		curl_setopt($ch, CURLOPT_REFERER, $input[$rand_keys]);
		curl_setopt($ch, CURLOPT_USERAGENT, $this->random_user_agent());
		curl_setopt($ch, CURLOPT_URL, $host);
		#curl_setopt($ch, CURLOPT_PROXYTYPE, CURLPROXY_HTTP);
		#curl_setopt($ch, CURLOPT_PROXY, $proxy);
		#curl_setopt($ch, CURLOPT_PROXYUSERPWD, $proxyauth);
		curl_setopt($ch, CURLOPT_TIMEOUT, 8500);
		curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
		curl_setopt($ch, CURLOPT_FOLLOWLOCATION, 1);
		$html = curl_exec($ch);
		echo curl_error($ch);
		curl_close($ch);
		return $html;
	} //end function curl_get

	/**
	 * @return mixed
	 */
	function random_user_agent()
	{
		$browser_freq = array("Internet Explorer" => 11.8, "Firefox" => 28.2, "Chrome" => 52.9, "Safari" => 3.9, "Opera" => 1.8);
		$browser_strings = array("Internet Explorer" => array("Mozilla/5.0 (compatible; MSIE 10.6; Windows NT 6.1; Trident/5.0; InfoPath.2; SLCC1; .NET CLR 3.0.4506.2152; .NET CLR 3.5.30729; .NET CLR 2.0.50727) 3gpp-gba UNTRUSTED/1.0", "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; WOW64; Trident/6.0)", "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; Trident/6.0)", "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; Trident/5.0)", "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; Trident/4.0; InfoPath.2; SV1; .NET CLR 2.0.50727; WOW64)", "Mozilla/5.0 (compatible; MSIE 10.0; Macintosh; Intel Mac OS X 10_7_3; Trident/6.0)", "Mozilla/4.0 (compatible; MSIE 10.0; Windows NT 6.1; Trident/5.0)", "Mozilla/1.22 (compatible; MSIE 10.0; Windows 3.1)", "Mozilla/5.0 (Windows; U; MSIE 9.0; WIndows NT 9.0; en-US))", "Mozilla/5.0 (Windows; U; MSIE 9.0; Windows NT 9.0; en-US)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 7.1; Trident/5.0)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0; SLCC2; Media Center PC 6.0; InfoPath.3; MS-RTC LM 8; Zune 4.7)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0; SLCC2; Media Center PC 6.0; InfoPath.3; MS-RTC LM 8; Zune 4.7", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; Zune 4.0; InfoPath.3; MS-RTC LM 8; .NET4.0C; .NET4.0E)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0; chromeframe/12.0.742.112)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0; .NET CLR 3.5.30729; .NET CLR 3.0.30729; .NET CLR 2.0.50727; Media Center PC 6.0)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Win64; x64; Trident/5.0; .NET CLR 3.5.30729; .NET CLR 3.0.30729; .NET CLR 2.0.50727; Media Center PC 6.0)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Win64; x64; Trident/5.0; .NET CLR 2.0.50727; SLCC2; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; Zune 4.0; Tablet PC 2.0; InfoPath.3; .NET4.0C; .NET4.0E)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Win64; x64; Trident/5.0", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0; yie8)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; InfoPath.2; .NET CLR 1.1.4322; .NET4.0C; Tablet PC 2.0)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0; FunWebProducts)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0; chromeframe/13.0.782.215)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0; chromeframe/11.0.696.57)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0) chromeframe/10.0.648.205", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/4.0; GTB7.4; InfoPath.1; SV1; .NET CLR 2.8.52393; WOW64; en-US)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.0; Trident/5.0; chromeframe/11.0.696.57)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.0; Trident/4.0; GTB7.4; InfoPath.3; SV1; .NET CLR 3.1.76908; WOW64; en-US)", "Mozilla/5.0 (Windows; U; MSIE 9.0; WIndows NT 9.0; en-US))", "Mozilla/5.0 (Windows; U; MSIE 9.0; Windows NT 9.0; en-US)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 7.1; Trident/5.0)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0; SLCC2; Media Center PC 6.0; InfoPath.3; MS-RTC LM 8; Zune 4.7)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0; SLCC2; Media Center PC 6.0; InfoPath.3; MS-RTC LM 8; Zune 4.7", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; Zune 4.0; InfoPath.3; MS-RTC LM 8; .NET4.0C; .NET4.0E)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0; chromeframe/12.0.742.112)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0; .NET CLR 3.5.30729; .NET CLR 3.0.30729; .NET CLR 2.0.50727; Media Center PC 6.0)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Win64; x64; Trident/5.0; .NET CLR 3.5.30729; .NET CLR 3.0.30729; .NET CLR 2.0.50727; Media Center PC 6.0)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Win64; x64; Trident/5.0; .NET CLR 2.0.50727; SLCC2; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; Zune 4.0; Tablet PC 2.0; InfoPath.3; .NET4.0C; .NET4.0E)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Win64; x64; Trident/5.0", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0; yie8)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; InfoPath.2; .NET CLR 1.1.4322; .NET4.0C; Tablet PC 2.0)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0; FunWebProducts)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0; chromeframe/13.0.782.215)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0; chromeframe/11.0.696.57)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0) chromeframe/10.0.648.205", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/4.0; GTB7.4; InfoPath.1; SV1; .NET CLR 2.8.52393; WOW64; en-US)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.0; Trident/5.0; chromeframe/11.0.696.57)", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.0; Trident/4.0; GTB7.4; InfoPath.3; SV1; .NET CLR 3.1.76908; WOW64; en-US)", "Mozilla/5.0 ( ; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)", "Mozilla/4.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/4.0; GTB7.4; InfoPath.2; SV1; .NET CLR 4.4.58799; WOW64; en-US)", "Mozilla/4.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/4.0; FDM; MSIECrawler; Media Center PC 5.0)", "Mozilla/4.0 (compatible; MSIE 9.0; Windows NT 6.0; Trident/4.0; GTB7.4; InfoPath.3; SV1; .NET CLR 3.4.53360; WOW64; en-US)", "Mozilla/4.0 (compatible; MSIE 9.0; Windows NT 5.1; Trident/5.0)", "Mozilla/4.0 (compatible; MSIE 9.0; Windows NT 5.1; Trident/4.0; .NET CLR 2.0.50727; .NET CLR 1.1.4322; .NET CLR 3.0.04506.648; .NET CLR 3.5.21022; .NET CLR 3.0.4506.2152; .NET CLR 3.5.30729; OfficeLiveConnector.1.4; OfficeLivePatch.1.3; .NET4.0C; .NE", "Mozilla/4.0 (compatible; MSIE 9.0; Windows 98; .NET CLR 3.0.04506.30)", "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 7.1; Trident/5.0; .NET CLR 2.0.50727; SLCC2; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; InfoPath.3; .NET4.0C)", "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/5.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; .NET4.0E; AskTB5.5)", "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; Trident/5.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; InfoPath.2; .NET4.0C; .NET4.0E)", "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; Win64; x64; Trident/5.0; .NET4.0C; .NET4.0E; InfoPath.3)", "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; Win64; x64; Trident/5.0; .NET CLR 2.0.50727; SLCC2; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; InfoPath.3; .NET4.0C)", "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; Trident/5.0; SLCC2; .NET CLR 2.0.50727; .NET4.0C)", "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; Trident/5.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; FDM; .NET CLR 1.1.4322; .NET4.0C; .NET4.0E; Tablet PC 2.0)", "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; Trident/5.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; Tablet PC 2.0; InfoPath.3; .NET4.0E)", "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0; Trident/5.0; SLCC1; .NET CLR 2.0.50727; Media Center PC 5.0; .NET CLR 3.5.30729; .NET CLR 3.0.30729; FDM; .NET4.0C; .NET4.0E; chromeframe/11.0.696.57)", "Mozilla/4.0 (compatible; U; MSIE 9.0; WIndows NT 9.0; en-US)", "Mozilla/4.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0; FunWebProducts)"), "Firefox" => array("Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:25.0) Gecko/20100101 Firefox/25.0", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.6; rv:25.0) Gecko/20100101 Firefox/25.0", "Mozilla/5.0 (Windows NT 6.0; WOW64; rv:24.0) Gecko/20100101 Firefox/24.0", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.8; rv:24.0) Gecko/20100101 Firefox/24.0", "Mozilla/5.0 (Windows NT 6.2; rv:22.0) Gecko/20130405 Firefox/23.0", "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:23.0) Gecko/20130406 Firefox/23.0", "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:23.0) Gecko/20131011 Firefox/23.0", "Mozilla/5.0 (Windows NT 6.2; rv:22.0) Gecko/20130405 Firefox/22.0", "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:22.0) Gecko/20130328 Firefox/22.0", "Mozilla/5.0 (Windows NT 6.1; rv:22.0) Gecko/20130405 Firefox/22.0", "Mozilla/5.0 (Windows NT 6.2; Win64; x64; rv:16.0.1) Gecko/20121011 Firefox/21.0.1", "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:16.0.1) Gecko/20121011 Firefox/21.0.1", "Mozilla/5.0 (Windows NT 6.2; Win64; x64; rv:21.0.0) Gecko/20121011 Firefox/21.0.0", "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:21.0) Gecko/20130331 Firefox/21.0", "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:21.0) Gecko/20100101 Firefox/21.0", "Mozilla/5.0 (X11; Linux i686; rv:21.0) Gecko/20100101 Firefox/21.0", "Mozilla/5.0 (Windows NT 6.2; rv:21.0) Gecko/20130326 Firefox/21.0", "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:21.0) Gecko/20130401 Firefox/21.0", "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:21.0) Gecko/20130331 Firefox/21.0", "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:21.0) Gecko/20130330 Firefox/21.0", "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:21.0) Gecko/20100101 Firefox/21.0", "Mozilla/5.0 (Windows NT 6.1; rv:21.0) Gecko/20130401 Firefox/21.0", "Mozilla/5.0 (Windows NT 6.1; rv:21.0) Gecko/20130328 Firefox/21.0", "Mozilla/5.0 (Windows NT 6.1; rv:21.0) Gecko/20100101 Firefox/21.0", "Mozilla/5.0 (Windows NT 5.1; rv:21.0) Gecko/20130401 Firefox/21.0", "Mozilla/5.0 (Windows NT 5.1; rv:21.0) Gecko/20130331 Firefox/21.0", "Mozilla/5.0 (Windows NT 5.1; rv:21.0) Gecko/20100101 Firefox/21.0", "Mozilla/5.0 (Windows NT 5.0; rv:21.0) Gecko/20100101 Firefox/21.0", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.8; rv:21.0) Gecko/20100101 Firefox/21.0", "Mozilla/5.0 (Windows NT 6.2; Win64; x64;) Gecko/20100101 Firefox/20.0", "Mozilla/5.0 (Windows NT 6.1; rv:6.0) Gecko/20100101 Firefox/19.0", "Mozilla/5.0 (Windows NT 6.1; rv:14.0) Gecko/20100101 Firefox/18.0.1", "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:18.0) Gecko/20100101 Firefox/18.0", "Mozilla/5.0 (X11; Ubuntu; Linux armv7l; rv:17.0) Gecko/20100101 Firefox/17.0", "Mozilla/6.0 (Windows NT 6.2; WOW64; rv:16.0.1) Gecko/20121011 Firefox/16.0.1", "Mozilla/5.0 (Windows NT 6.2; WOW64; rv:16.0.1) Gecko/20121011 Firefox/16.0.1", "Mozilla/5.0 (Windows NT 6.2; Win64; x64; rv:16.0.1) Gecko/20121011 Firefox/16.0.1", "Mozilla/5.0 (Windows NT 6.1; rv:15.0) Gecko/20120716 Firefox/15.0a2", "Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.9.1.16) Gecko/20120427 Firefox/15.0a1", "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:15.0) Gecko/20120427 Firefox/15.0a1", "Mozilla/5.0 (Windows NT 6.2; WOW64; rv:15.0) Gecko/20120910144328 Firefox/15.0.2", "Mozilla/5.0 (X11; Ubuntu; Linux i686; rv:15.0) Gecko/20100101 Firefox/15.0.1", "Mozilla/5.0 (Windows; U; Windows NT 5.1; rv:15.0) Gecko/20121011 Firefox/15.0.1"), "Chrome" => array("Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/30.0.1599.17 Safari/537.36", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/29.0.1547.62 Safari/537.36", "Mozilla/5.0 (X11; CrOS i686 4319.74.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/29.0.1547.57 Safari/537.36", "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/29.0.1547.2 Safari/537.36", "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/28.0.1468.0 Safari/537.36", "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/28.0.1467.0 Safari/537.36", "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/28.0.1464.0 Safari/537.36", "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/27.0.1453.93 Safari/537.36", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/27.0.1453.93 Safari/537.36", "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/27.0.1453.93 Safari/537.36", "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/27.0.1453.93 Safari/537.36", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_8_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/27.0.1453.93 Safari/537.36", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/27.0.1453.93 Safari/537.36", "Mozilla/5.0 (X11; CrOS i686 3912.101.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/27.0.1453.116 Safari/537.36", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.17 (KHTML, like Gecko) Chrome/24.0.1312.60 Safari/537.17", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_8_2) AppleWebKit/537.17 (KHTML, like Gecko) Chrome/24.0.1309.0 Safari/537.17", "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.15 (KHTML, like Gecko) Chrome/24.0.1295.0 Safari/537.15", "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.14 (KHTML, like Gecko) Chrome/24.0.1292.0 Safari/537.14", "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.13 (KHTML, like Gecko) Chrome/24.0.1290.1 Safari/537.13", "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/537.13 (KHTML, like Gecko) Chrome/24.0.1290.1 Safari/537.13", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_8_2) AppleWebKit/537.13 (KHTML, like Gecko) Chrome/24.0.1290.1 Safari/537.13", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_4) AppleWebKit/537.13 (KHTML, like Gecko) Chrome/24.0.1290.1 Safari/537.13", "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.13 (KHTML, like Gecko) Chrome/24.0.1284.0 Safari/537.13", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_8_2) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.6 Safari/537.11", "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.26 Safari/537.11", "Mozilla/5.0 (Windows NT 6.0) yi; AppleWebKit/345667.12221 (KHTML, like Gecko) Chrome/23.0.1271.26 Safari/453667.1221", "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.11 (KHTML, like Gecko) Chrome/23.0.1271.17 Safari/537.11", "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/537.4 (KHTML, like Gecko) Chrome/22.0.1229.94 Safari/537.4", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_6_0) AppleWebKit/537.4 (KHTML, like Gecko) Chrome/22.0.1229.79 Safari/537.4", "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.2 (KHTML, like Gecko) Chrome/22.0.1216.0 Safari/537.2", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/22.0.1207.1 Safari/537.1", "Mozilla/5.0 (X11; CrOS i686 2268.111.0) AppleWebKit/536.11 (KHTML, like Gecko) Chrome/20.0.1132.57 Safari/536.11", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.6 (KHTML, like Gecko) Chrome/20.0.1092.0 Safari/536.6", "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.6 (KHTML, like Gecko) Chrome/20.0.1090.0 Safari/536.6", "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/537.1 (KHTML, like Gecko) Chrome/19.77.34.5 Safari/537.1", "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.9 Safari/536.5", "Mozilla/5.0 (Windows NT 6.0) AppleWebKit/536.5 (KHTML, like Gecko) Chrome/19.0.1084.36 Safari/536.5", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1063.0 Safari/536.3", "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1063.0 Safari/536.3", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_8_0) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1063.0 Safari/536.3", "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1062.0 Safari/536.3", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1062.0 Safari/536.3", "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.1 Safari/536.3", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.1 Safari/536.3", "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.1 Safari/536.3", "Mozilla/5.0 (Windows NT 6.2) AppleWebKit/536.3 (KHTML, like Gecko) Chrome/19.0.1061.0 Safari/536.3", "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/535.24 (KHTML, like Gecko) Chrome/19.0.1055.1 Safari/535.24", "Mozilla/5.0 (Windows NT 6.2; WOW64) AppleWebKit/535.24 (KHTML, like Gecko) Chrome/19.0.1055.1 Safari/535.24", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_2) AppleWebKit/535.24 (KHTML, like Gecko) Chrome/19.0.1055.1 Safari/535.24", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_3) AppleWebKit/535.22 (KHTML, like Gecko) Chrome/19.0.1047.0 Safari/535.22", "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/535.21 (KHTML, like Gecko) Chrome/19.0.1042.0 Safari/535.21", "Mozilla/5.0 (X11; Linux i686) AppleWebKit/535.21 (KHTML, like Gecko) Chrome/19.0.1041.0 Safari/535.21", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_3) AppleWebKit/535.20 (KHTML, like Gecko) Chrome/19.0.1036.7 Safari/535.20", "Mozilla/5.0 (Windows NT 6.1) AppleWebKit/535.2 (KHTML, like Gecko) Chrome/18.6.872.0 Safari/535.2 UNTRUSTED/1.0 3gpp-gba UNTRUSTED/1.0", "Mozilla/5.0 (Macintosh; AMD Mac OS X 10_8_2) AppleWebKit/535.22 (KHTML, like Gecko) Chrome/18.6.872", "Mozilla/5.0 (X11; CrOS i686 1660.57.0) AppleWebKit/535.19 (KHTML, like Gecko) Chrome/18.0.1025.46 Safari/535.19", "Mozilla/5.0 (Windows NT 6.0; WOW64) AppleWebKit/535.19 (KHTML, like Gecko) Chrome/18.0.1025.45 Safari/535.19", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_2) AppleWebKit/535.19 (KHTML, like Gecko) Chrome/18.0.1025.45 Safari/535.19", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_6_8) AppleWebKit/535.19 (KHTML, like Gecko) Chrome/18.0.1025.45 Safari/535.19"), "Safari" => array("Mozilla/5.0 (iPad; CPU OS 6_0 like Mac OS X) AppleWebKit/536.26 (KHTML, like Gecko) Version/6.0 Mobile/10A5355d Safari/8536.25", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_6_8) AppleWebKit/537.13+ (KHTML, like Gecko) Version/5.1.7 Safari/534.57.2", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_3) AppleWebKit/534.55.3 (KHTML, like Gecko) Version/5.1.3 Safari/534.53.10", "Mozilla/5.0 (iPad; CPU OS 5_1 like Mac OS X) AppleWebKit/534.46 (KHTML, like Gecko ) Version/5.1 Mobile/9B176 Safari/7534.48.3", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_8; de-at) AppleWebKit/533.21.1 (KHTML, like Gecko) Version/5.0.5 Safari/533.21.1", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_7; da-dk) AppleWebKit/533.21.1 (KHTML, like Gecko) Version/5.0.5 Safari/533.21.1", "Mozilla/5.0 (Windows; U; Windows NT 6.1; tr-TR) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.4 Safari/533.20.27", "Mozilla/5.0 (Windows; U; Windows NT 6.1; ko-KR) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.4 Safari/533.20.27", "Mozilla/5.0 (Windows; U; Windows NT 6.1; fr-FR) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.4 Safari/533.20.27", "Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.4 Safari/533.20.27", "Mozilla/5.0 (Windows; U; Windows NT 6.1; cs-CZ) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.4 Safari/533.20.27", "Mozilla/5.0 (Windows; U; Windows NT 6.0; ja-JP) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.4 Safari/533.20.27", "Mozilla/5.0 (Windows; U; Windows NT 6.0; en-US) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.4 Safari/533.20.27", "Mozilla/5.0 (Macintosh; U; PPC Mac OS X 10_5_8; zh-cn) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.4 Safari/533.20.27", "Mozilla/5.0 (Macintosh; U; PPC Mac OS X 10_5_8; ja-jp) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.4 Safari/533.20.27", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_7; ja-jp) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.4 Safari/533.20.27", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_6; zh-cn) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.4 Safari/533.20.27", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_6; sv-se) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.4 Safari/533.20.27", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_6; ko-kr) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.4 Safari/533.20.27", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_6; ja-jp) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.4 Safari/533.20.27", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_6; it-it) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.4 Safari/533.20.27", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_6; fr-fr) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.4 Safari/533.20.27", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_6; es-es) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.4 Safari/533.20.27", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_6; en-us) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.4 Safari/533.20.27", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_6; en-gb) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.4 Safari/533.20.27", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_6; de-de) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.4 Safari/533.20.27", "Mozilla/5.0 (Windows; U; Windows NT 6.1; sv-SE) AppleWebKit/533.19.4 (KHTML, like Gecko) Version/5.0.3 Safari/533.19.4", "Mozilla/5.0 (Windows; U; Windows NT 6.1; ja-JP) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.3 Safari/533.19.4", "Mozilla/5.0 (Windows; U; Windows NT 6.1; de-DE) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.3 Safari/533.19.4", "Mozilla/5.0 (Windows; U; Windows NT 6.0; hu-HU) AppleWebKit/533.19.4 (KHTML, like Gecko) Version/5.0.3 Safari/533.19.4", "Mozilla/5.0 (Windows; U; Windows NT 6.0; en-US) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.3 Safari/533.19.4", "Mozilla/5.0 (Windows; U; Windows NT 6.0; de-DE) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.3 Safari/533.19.4", "Mozilla/5.0 (Windows; U; Windows NT 5.1; ru-RU) AppleWebKit/533.19.4 (KHTML, like Gecko) Version/5.0.3 Safari/533.19.4", "Mozilla/5.0 (Windows; U; Windows NT 5.1; ja-JP) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.3 Safari/533.19.4", "Mozilla/5.0 (Windows; U; Windows NT 5.1; it-IT) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.3 Safari/533.19.4", "Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US) AppleWebKit/533.20.25 (KHTML, like Gecko) Version/5.0.3 Safari/533.19.4", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_7; en-us) AppleWebKit/534.16+ (KHTML, like Gecko) Version/5.0.3 Safari/533.19.4", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_6; fr-ch) AppleWebKit/533.19.4 (KHTML, like Gecko) Version/5.0.3 Safari/533.19.4", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_5; de-de) AppleWebKit/534.15+ (KHTML, like Gecko) Version/5.0.3 Safari/533.19.4", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_5; ar) AppleWebKit/533.19.4 (KHTML, like Gecko) Version/5.0.3 Safari/533.19.4", "Mozilla/5.0 (Android 2.2; Windows; U; Windows NT 6.1; en-US) AppleWebKit/533.19.4 (KHTML, like Gecko) Version/5.0.3 Safari/533.19.4", "Mozilla/5.0 (Windows; U; Windows NT 6.1; zh-HK) AppleWebKit/533.18.1 (KHTML, like Gecko) Version/5.0.2 Safari/533.18.5", "Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US) AppleWebKit/533.19.4 (KHTML, like Gecko) Version/5.0.2 Safari/533.18.5", "Mozilla/5.0 (Windows; U; Windows NT 6.0; tr-TR) AppleWebKit/533.18.1 (KHTML, like Gecko) Version/5.0.2 Safari/533.18.5", "Mozilla/5.0 (Windows; U; Windows NT 6.0; nb-NO) AppleWebKit/533.18.1 (KHTML, like Gecko) Version/5.0.2 Safari/533.18.5", "Mozilla/5.0 (Windows; U; Windows NT 6.0; fr-FR) AppleWebKit/533.18.1 (KHTML, like Gecko) Version/5.0.2 Safari/533.18.5", "Mozilla/5.0 (Windows; U; Windows NT 5.1; zh-TW) AppleWebKit/533.19.4 (KHTML, like Gecko) Version/5.0.2 Safari/533.18.5", "Mozilla/5.0 (Windows; U; Windows NT 5.1; ru-RU) AppleWebKit/533.18.1 (KHTML, like Gecko) Version/5.0.2 Safari/533.18.5", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_5_8; zh-cn) AppleWebKit/533.18.1 (KHTML, like Gecko) Version/5.0.2 Safari/533.18.5", "Mozilla/5.0 (iPod; U; CPU iPhone OS 4_3_3 like Mac OS X; ja-jp) AppleWebKit/533.17.9 (KHTML, like Gecko) Version/5.0.2 Mobile/8J2 Safari/6533.18.5", "Mozilla/5.0 (iPod; U; CPU iPhone OS 4_3_1 like Mac OS X; zh-cn) AppleWebKit/533.17.9 (KHTML, like Gecko) Version/5.0.2 Mobile/8G4 Safari/6533.18.5", "Mozilla/5.0 (iPod; U; CPU iPhone OS 4_2_1 like Mac OS X; he-il) AppleWebKit/533.17.9 (KHTML, like Gecko) Version/5.0.2 Mobile/8C148 Safari/6533.18.5", "Mozilla/5.0 (iPhone; U; ru; CPU iPhone OS 4_2_1 like Mac OS X; ru) AppleWebKit/533.17.9 (KHTML, like Gecko) Version/5.0.2 Mobile/8C148a Safari/6533.18.5", "Mozilla/5.0 (iPhone; U; ru; CPU iPhone OS 4_2_1 like Mac OS X; fr) AppleWebKit/533.17.9 (KHTML, like Gecko) Version/5.0.2 Mobile/8C148a Safari/6533.18.5", "Mozilla/5.0 (iPhone; U; fr; CPU iPhone OS 4_2_1 like Mac OS X; fr) AppleWebKit/533.17.9 (KHTML, like Gecko) Version/5.0.2 Mobile/8C148a Safari/6533.18.5", "Mozilla/5.0 (iPhone; U; CPU iPhone OS 4_3_1 like Mac OS X; zh-tw) AppleWebKit/533.17.9 (KHTML, like Gecko) Version/5.0.2 Mobile/8G4 Safari/6533.18.5", "Mozilla/5.0 (iPhone; U; CPU iPhone OS 4_3 like Mac OS X; pl-pl) AppleWebKit/533.17.9 (KHTML, like Gecko) Version/5.0.2 Mobile/8F190 Safari/6533.18.5", "Mozilla/5.0 (iPhone; U; CPU iPhone OS 4_3 like Mac OS X; fr-fr) AppleWebKit/533.17.9 (KHTML, like Gecko) Version/5.0.2 Mobile/8F190 Safari/6533.18.5", "Mozilla/5.0 (iPhone; U; CPU iPhone OS 4_3 like Mac OS X; en-gb) AppleWebKit/533.17.9 (KHTML, like Gecko) Version/5.0.2 Mobile/8F190 Safari/6533.18.5", "Mozilla/5.0 (iPhone; U; CPU iPhone OS 4_2_1 like Mac OS X; ru-ru) AppleWebKit/533.17.9 (KHTML, like Gecko) Version/5.0.2 Mobile/8C148 Safari/6533.18.5", "Mozilla/5.0 (iPhone; U; CPU iPhone OS 4_2_1 like Mac OS X; nb-no) AppleWebKit/533.17.9 (KHTML, like Gecko) Version/5.0.2 Mobile/8C148a Safari/6533.18.5", "Mozilla/5.0 (Windows; U; Windows NT 5.2; en-US) AppleWebKit/533.17.8 (KHTML, like Gecko) Version/5.0.1 Safari/533.17.8", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_4; th-th) AppleWebKit/533.17.8 (KHTML, like Gecko) Version/5.0.1 Safari/533.17.8", "Mozilla/5.0 (X11; U; Linux x86_64; en-us) AppleWebKit/531.2+ (KHTML, like Gecko) Version/5.0 Safari/531.2+", "Mozilla/5.0 (X11; U; Linux x86_64; en-ca) AppleWebKit/531.2+ (KHTML, like Gecko) Version/5.0 Safari/531.2+", "Mozilla/5.0 (Windows; U; Windows NT 6.1; ja-JP) AppleWebKit/533.16 (KHTML, like Gecko) Version/5.0 Safari/533.16", "Mozilla/5.0 (Windows; U; Windows NT 6.1; es-ES) AppleWebKit/533.18.1 (KHTML, like Gecko) Version/5.0 Safari/533.16", "Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US) AppleWebKit/533.18.1 (KHTML, like Gecko) Version/5.0 Safari/533.16", "Mozilla/5.0 (Windows; U; Windows NT 6.0; ja-JP) AppleWebKit/533.16 (KHTML, like Gecko) Version/5.0 Safari/533.16", "Mozilla/5.0 (Macintosh; U; PPC Mac OS X 10_5_8; ja-jp) AppleWebKit/533.16 (KHTML, like Gecko) Version/5.0 Safari/533.16", "Mozilla/5.0 (Macintosh; U; PPC Mac OS X 10_4_11; fr) AppleWebKit/533.16 (KHTML, like Gecko) Version/5.0 Safari/533.16", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_3; zh-cn) AppleWebKit/533.16 (KHTML, like Gecko) Version/5.0 Safari/533.16", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_3; ru-ru) AppleWebKit/533.16 (KHTML, like Gecko) Version/5.0 Safari/533.16", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_3; ko-kr) AppleWebKit/533.16 (KHTML, like Gecko) Version/5.0 Safari/533.16", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_3; it-it) AppleWebKit/533.16 (KHTML, like Gecko) Version/5.0 Safari/533.16", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_3; HTC-P715a; en-ca) AppleWebKit/533.16 (KHTML, like Gecko) Version/5.0 Safari/533.16", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_3; en-us) AppleWebKit/534.1+ (KHTML, like Gecko) Version/5.0 Safari/533.16", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_3; en-au) AppleWebKit/533.16 (KHTML, like Gecko) Version/5.0 Safari/533.16", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_3; el-gr) AppleWebKit/533.16 (KHTML, like Gecko) Version/5.0 Safari/533.16", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_3; ca-es) AppleWebKit/533.16 (KHTML, like Gecko) Version/5.0 Safari/533.16", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_5_8; zh-tw) AppleWebKit/533.16 (KHTML, like Gecko) Version/5.0 Safari/533.16", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_5_8; ja-jp) AppleWebKit/533.16 (KHTML, like Gecko) Version/5.0 Safari/533.16", "Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_5_8; it-it) AppleWebKit/533.16 (KHTML, like Gecko) Version/5.0 Safari/533.16"), "Opera" => array("Opera/9.80 (Windows NT 6.0) Presto/2.12.388 Version/12.14", "Mozilla/5.0 (Windows NT 6.0; rv:2.0) Gecko/20100101 Firefox/4.0 Opera 12.14", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.0) Opera 12.14", "Opera/12.80 (Windows NT 5.1; U; en) Presto/2.10.289 Version/12.02", "Opera/9.80 (Windows NT 6.1; U; es-ES) Presto/2.9.181 Version/12.00", "Opera/9.80 (Windows NT 5.1; U; zh-sg) Presto/2.9.181 Version/12.00", "Opera/12.0(Windows NT 5.2;U;en)Presto/22.9.168 Version/12.00", "Opera/12.0(Windows NT 5.1;U;en)Presto/22.9.168 Version/12.00", "Mozilla/5.0 (Windows NT 5.1) Gecko/20100101 Firefox/14.0 Opera/12.0", "Opera/9.80 (Windows NT 6.1; WOW64; U; pt) Presto/2.10.229 Version/11.62", "Opera/9.80 (Windows NT 6.0; U; pl) Presto/2.10.229 Version/11.62", "Opera/9.80 (Macintosh; Intel Mac OS X 10.6.8; U; fr) Presto/2.9.168 Version/11.52", "Opera/9.80 (Macintosh; Intel Mac OS X 10.6.8; U; de) Presto/2.9.168 Version/11.52", "Opera/9.80 (Windows NT 5.1; U; en) Presto/2.9.168 Version/11.51", "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; de) Opera 11.51", "Opera/9.80 (X11; Linux x86_64; U; fr) Presto/2.9.168 Version/11.50", "Opera/9.80 (X11; Linux i686; U; hu) Presto/2.9.168 Version/11.50", "Opera/9.80 (X11; Linux i686; U; ru) Presto/2.8.131 Version/11.11", "Opera/9.80 (X11; Linux i686; U; es-ES) Presto/2.8.131 Version/11.11", "Mozilla/5.0 (Windows NT 5.1; U; en; rv:1.8.1) Gecko/20061208 Firefox/5.0 Opera 11.11", "Opera/9.80 (X11; Linux x86_64; U; bg) Presto/2.8.131 Version/11.10", "Opera/9.80 (Windows NT 6.0; U; en) Presto/2.8.99 Version/11.10", "Opera/9.80 (Windows NT 5.1; U; zh-tw) Presto/2.8.131 Version/11.10", "Opera/9.80 (Windows NT 6.1; Opera Tablet/15165; U; en) Presto/2.8.149 Version/11.1", "Opera/9.80 (X11; Linux x86_64; U; Ubuntu/10.10 (maverick); pl) Presto/2.7.62 Version/11.01", "Opera/9.80 (X11; Linux i686; U; ja) Presto/2.7.62 Version/11.01", "Opera/9.80 (X11; Linux i686; U; fr) Presto/2.7.62 Version/11.01", "Opera/9.80 (Windows NT 6.1; U; zh-tw) Presto/2.7.62 Version/11.01", "Opera/9.80 (Windows NT 6.1; U; zh-cn) Presto/2.7.62 Version/11.01", "Opera/9.80 (Windows NT 6.1; U; sv) Presto/2.7.62 Version/11.01", "Opera/9.80 (Windows NT 6.1; U; en-US) Presto/2.7.62 Version/11.01", "Opera/9.80 (Windows NT 6.1; U; cs) Presto/2.7.62 Version/11.01", "Opera/9.80 (Windows NT 6.0; U; pl) Presto/2.7.62 Version/11.01", "Opera/9.80 (Windows NT 5.2; U; ru) Presto/2.7.62 Version/11.01", "Opera/9.80 (Windows NT 5.1; U;) Presto/2.7.62 Version/11.01", "Opera/9.80 (Windows NT 5.1; U; cs) Presto/2.7.62 Version/11.01", "Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US; rv:1.9.2.13) Gecko/20101213 Opera/9.80 (Windows NT 6.1; U; zh-tw) Presto/2.7.62 Version/11.01", "Mozilla/5.0 (Windows NT 6.1; U; nl; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6 Opera 11.01", "Mozilla/5.0 (Windows NT 6.1; U; de; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6 Opera 11.01", "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; de) Opera 11.01", "Opera/9.80 (X11; Linux x86_64; U; pl) Presto/2.7.62 Version/11.00", "Opera/9.80 (X11; Linux i686; U; it) Presto/2.7.62 Version/11.00", "Opera/9.80 (Windows NT 6.1; U; zh-cn) Presto/2.6.37 Version/11.00", "Opera/9.80 (Windows NT 6.1; U; pl) Presto/2.7.62 Version/11.00", "Opera/9.80 (Windows NT 6.1; U; ko) Presto/2.7.62 Version/11.00", "Opera/9.80 (Windows NT 6.1; U; fi) Presto/2.7.62 Version/11.00", "Opera/9.80 (Windows NT 6.1; U; en-GB) Presto/2.7.62 Version/11.00", "Opera/9.80 (Windows NT 6.1 x64; U; en) Presto/2.7.62 Version/11.00", "Opera/9.80 (Windows NT 6.0; U; en) Presto/2.7.39 Version/11.00", "Opera/9.80 (Windows NT 5.1; U; ru) Presto/2.7.39 Version/11.00", "Opera/9.80 (Windows NT 5.1; U; MRA 5.5 (build 02842); ru) Presto/2.7.62 Version/11.00", "Opera/9.80 (Windows NT 5.1; U; it) Presto/2.7.62 Version/11.00", "Mozilla/5.0 (Windows NT 6.0; U; ja; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6 Opera 11.00", "Mozilla/5.0 (Windows NT 5.1; U; pl; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6 Opera 11.00", "Mozilla/5.0 (Windows NT 5.1; U; de; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6 Opera 11.00", "Mozilla/4.0 (compatible; MSIE 8.0; X11; Linux x86_64; pl) Opera 11.00", "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; fr) Opera 11.00", "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; ja) Opera 11.00", "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; en) Opera 11.00", "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 5.1; pl) Opera 11.00", "Opera/9.80 (Windows NT 6.1; U; pl) Presto/2.6.31 Version/10.70", "Mozilla/5.0 (Windows NT 5.2; U; ru; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6 Opera 10.70", "Mozilla/5.0 (Windows NT 5.1; U; zh-cn; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6 Opera 10.70", "Opera/9.80 (Windows NT 5.2; U; zh-cn) Presto/2.6.30 Version/10.63", "Opera/9.80 (Windows NT 5.2; U; en) Presto/2.6.30 Version/10.63", "Opera/9.80 (Windows NT 5.1; U; MRA 5.6 (build 03278); ru) Presto/2.6.30 Version/10.63", "Opera/9.80 (Windows NT 5.1; U; pl) Presto/2.6.30 Version/10.62", "Mozilla/5.0 (X11; Linux x86_64; U; de; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6 Opera 10.62", "Mozilla/4.0 (compatible; MSIE 8.0; X11; Linux x86_64; de) Opera 10.62", "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; en) Opera 10.62", "Opera/9.80 (X11; Linux i686; U; pl) Presto/2.6.30 Version/10.61", "Opera/9.80 (X11; Linux i686; U; es-ES) Presto/2.6.30 Version/10.61", "Opera/9.80 (Windows NT 6.1; U; zh-cn) Presto/2.6.30 Version/10.61", "Opera/9.80 (Windows NT 6.1; U; en) Presto/2.6.30 Version/10.61", "Opera/9.80 (Windows NT 6.0; U; it) Presto/2.6.30 Version/10.61", "Opera/9.80 (Windows NT 5.2; U; ru) Presto/2.6.30 Version/10.61", "Opera/9.80 (Windows 98; U; de) Presto/2.6.30 Version/10.61", "Opera/9.80 (Macintosh; Intel Mac OS X; U; nl) Presto/2.6.30 Version/10.61", "Opera/9.80 (X11; Linux i686; U; en) Presto/2.5.27 Version/10.60", "Opera/9.80 (Windows NT 6.0; U; nl) Presto/2.6.30 Version/10.60", "Opera/10.60 (Windows NT 5.1; U; zh-cn) Presto/2.6.30 Version/10.60", "Opera/10.60 (Windows NT 5.1; U; en-US) Presto/2.6.30 Version/10.60", "Opera/9.80 (X11; Linux i686; U; it) Presto/2.5.24 Version/10.54", "Opera/9.80 (X11; Linux i686; U; en-GB) Presto/2.5.24 Version/10.53", "Mozilla/5.0 (Windows NT 5.1; U; zh-cn; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6 Opera 10.53", "Mozilla/5.0 (Windows NT 5.1; U; Firefox/5.0; en; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6 Opera 10.53", "Mozilla/5.0 (Windows NT 5.1; U; Firefox/4.5; en; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6 Opera 10.53", "Mozilla/5.0 (Windows NT 5.1; U; Firefox/3.5; en; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6 Opera 10.53", "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 5.1; ko) Opera 10.53", "Opera/9.80 (Windows NT 6.1; U; fr) Presto/2.5.24 Version/10.52", "Opera/9.80 (Windows NT 6.1; U; en) Presto/2.5.22 Version/10.51", "Opera/9.80 (Windows NT 6.0; U; cs) Presto/2.5.22 Version/10.51", "Opera/9.80 (Windows NT 5.2; U; ru) Presto/2.5.22 Version/10.51", "Opera/9.80 (Linux i686; U; en) Presto/2.5.22 Version/10.51", "Mozilla/5.0 (Windows NT 6.1; U; en-GB; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6 Opera 10.51", "Mozilla/5.0 (Linux i686; U; en; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6 Opera 10.51", "Mozilla/4.0 (compatible; MSIE 8.0; Linux i686; en) Opera 10.51", "Opera/9.80 (Windows NT 6.1; U; zh-tw) Presto/2.5.22 Version/10.50", "Opera/9.80 (Windows NT 6.1; U; zh-cn) Presto/2.5.22 Version/10.50", "Opera/9.80 (Windows NT 6.1; U; sk) Presto/2.6.22 Version/10.50", "Opera/9.80 (Windows NT 6.1; U; ja) Presto/2.5.22 Version/10.50", "Opera/9.80 (Windows NT 6.0; U; zh-cn) Presto/2.5.22 Version/10.50", "Opera/9.80 (Windows NT 5.1; U; sk) Presto/2.5.22 Version/10.50", "Opera/9.80 (Windows NT 5.1; U; ru) Presto/2.5.22 Version/10.50", "Opera/10.50 (Windows NT 6.1; U; en-GB) Presto/2.2.2", "Opera/9.80 (S60; SymbOS; Opera Tablet/9174; U; en) Presto/2.7.81 Version/10.5", "Opera/9.80 (X11; U; Linux i686; en-US; rv:1.9.2.3) Presto/2.2.15 Version/10.10", "Opera/9.80 (X11; Linux x86_64; U; it) Presto/2.2.15 Version/10.10", "Opera/9.80 (Windows NT 6.1; U; de) Presto/2.2.15 Version/10.10", "Opera/9.80 (Windows NT 6.0; U; Gecko/20100115; pl) Presto/2.2.15 Version/10.10", "Opera/9.80 (Windows NT 6.0; U; en) Presto/2.2.15 Version/10.10", "Opera/9.80 (Windows NT 5.1; U; de) Presto/2.2.15 Version/10.10", "Opera/9.80 (Windows NT 5.1; U; cs) Presto/2.2.15 Version/10.10", "Mozilla/5.0 (Windows NT 6.0; U; tr; rv:1.8.1) Gecko/20061208 Firefox/2.0.0 Opera 10.10", "Mozilla/4.0 (compatible; MSIE 6.0; X11; Linux i686; de) Opera 10.10", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 6.0; tr) Opera 10.10", "Opera/9.80 (X11; Linux x86_64; U; en-GB) Presto/2.2.15 Version/10.01", "Opera/9.80 (X11; Linux x86_64; U; en) Presto/2.2.15 Version/10.00", "Opera/9.80 (X11; Linux x86_64; U; de) Presto/2.2.15 Version/10.00", "Opera/9.80 (X11; Linux i686; U; ru) Presto/2.2.15 Version/10.00", "Opera/9.80 (X11; Linux i686; U; pt-BR) Presto/2.2.15 Version/10.00", "Opera/9.80 (X11; Linux i686; U; pl) Presto/2.2.15 Version/10.00", "Opera/9.80 (X11; Linux i686; U; nb) Presto/2.2.15 Version/10.00", "Opera/9.80 (X11; Linux i686; U; en-GB) Presto/2.2.15 Version/10.00", "Opera/9.80 (X11; Linux i686; U; en) Presto/2.2.15 Version/10.00", "Opera/9.80 (X11; Linux i686; U; Debian; pl) Presto/2.2.15 Version/10.00", "Opera/9.80 (X11; Linux i686; U; de) Presto/2.2.15 Version/10.00", "Opera/9.80 (Windows NT 6.1; U; zh-cn) Presto/2.2.15 Version/10.00", "Opera/9.80 (Windows NT 6.1; U; fi) Presto/2.2.15 Version/10.00", "Opera/9.80 (Windows NT 6.1; U; en) Presto/2.2.15 Version/10.00", "Opera/9.80 (Windows NT 6.1; U; de) Presto/2.2.15 Version/10.00", "Opera/9.80 (Windows NT 6.1; U; cs) Presto/2.2.15 Version/10.00", "Opera/9.80 (Windows NT 6.0; U; en) Presto/2.2.15 Version/10.00", "Opera/9.80 (Windows NT 6.0; U; de) Presto/2.2.15 Version/10.00", "Opera/9.80 (Windows NT 5.2; U; en) Presto/2.2.15 Version/10.00", "Opera/9.80 (Windows NT 5.1; U; zh-cn) Presto/2.2.15 Version/10.00", "Opera/9.80 (Windows NT 5.1; U; ru) Presto/2.2.15 Version/10.00"));

		$max = 0;
		$rcount = 0;
		$browser_type = '';

		foreach ($browser_freq as $k => $v) $max += $v;
		$roll = rand(0, $max);
		foreach ($browser_freq as $k => $v) if(($roll <= ($rcount += $v)) and (!$browser_type)) $browser_type = $k;
		$user_agent_array = $browser_strings[$browser_type];
		shuffle($user_agent_array);
		$user_agent = $user_agent_array[0];
		return $user_agent;
	}

	/**
	 * @param $res
	 *
	 * выолняет множественный запуск функций парсинга
	 */
	function showShopContent($res)
	{
		foreach ($res as $key => $val) {
        ?>
            <a class="spoiler" href="#spoiler" tabindex="1">
            <? echo  $key ;?>
            <div class="spoiler-content">
        <?
			foreach ($val as $item) {
				if($item[0] != '') {
                    if($item[0] != 'нет')
                    {
                        echo $item[0] . " р. - " . $item[1] . "<br>";
                    }
                    else
                    {
                        echo $item[0] . " - " . $item[1] . "<br>";
                    }
				}
			}?>
            </div>
        </a>
		<?}
	}
}

?>