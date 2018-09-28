<?php

require_once ('validateEmail.php');

$emails = array('aleksejt@scandiweb.com', 'aleksejnano@scandiweb.com', 'al.tseb91@gmail.com', 'aloha.tseboha@gmail.com');

$sender = "leksa.ukr@gmail.com";

$smtpValidator = new validateEmail();

$result = $smtpValidator->validate($emails, $sender);

$something = "a";