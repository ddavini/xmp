<?php

/*  *****************************************************************************
**
**  RC4Crypt 3.0
**
**  2004/11/23 -- Correct naming, new way of calling RC4Class
**  2004/04/28 -- Bug fix for incorrect output due to XORing stupidity
**                [Ariel Arjona (beerfrick@users.sourceforge.net)]
**  2003/11/15 -- Cleans urlencode's, fixed the invalid outputs
**  2003/02/22 -- A fix, which turned out to be a bug!
**  2001/02/24 -- Passes RC4 Vector Harness
**
**  Website     : http://www.devhome.org
**  Email       : god@mjsabby.com
**
**  Copyright Notice:
**
**  RC4 is a registered trademark of the RSA Data Security Inc.
**  This is an implementation of the original algorithm. The author
**  of this program is NOT the original publisher of this algorithm.
**
**  (C) Copyright 2003, 2004 Mukul Sabharwal [http://mjsabby.com]
**  All Rights Reserved
**
**  ****************************************************************************/

/**
 * This program is free software; you can redistribute it and/or
 * modify it under the terms of the GNU General Public License
 * as published by the Free Software Foundation; either version 2
 * of the License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * The GNU General Public License can be found at
 * http://www.gnu.org/copyleft/gpl.html
 *
 * You should have received a copy of the GNU General Public License
 * along with this program; if not, write to the Free Software
 * Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
 *
 * @link http://www.devhome.org/php/tutorials/rc4crypt.html
 * @version 3.0
 * @copyright Copyright 2003, 2004 Mukul Sabharwal
 * @author Mukul Sabharwal <god@mjsabby.com>
 * @access public
 * @package RC4Crypt
 */

/* $Id: rc4test.php,v 1.1 2006/01/11 23:12:07 x-mad Exp $ */

include('./class.rc4crypt.php');

if (strcmp('4.1.0', phpversion()) > 0) {
	global $HTTP_GET_VARS;
	$_GET = &$HTTP_GET_VARS;
}

$pwd = $_GET['pwd'];
$data = $_GET['data'];

if (get_magic_quotes_gpc()) {
	$pwd = stripslashes($pwd);
	$data = stripslashes($data);
}

$e = rc4crypt::encrypt($pwd, $data);
$d = rc4crypt::decrypt($pwd, $e);

echo $d;
echo $e;

?>