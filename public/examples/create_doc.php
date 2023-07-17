<?php

require_once '../../vendor/autoload.php';

# use the factory to create a Faker\Generator instance
$faker = Faker\Factory::create("pt_BR");

$users = array();

for ($i = 0; $i < 50; $i++) {

    $users[] = array(
        "id" => $i + 1,
        "name" => $faker->name(),
        "cpf" => fakeCpfGeneration(),
        "phone_number" => $faker->phoneNumber(),
        "email" => $faker->email(),
        "state" => $faker->state()
    );
}

$header = [
    "#",
    "Nome",
    "cpf",
    "telefone",
    "email",
    "Estado representado",
];

try {

    (new \App\Spreadsheets())->writeWorksheet(["to_style_stripes" => true], $header, $users);

} catch (\Exception $e) {
    echo $e->getMessage() . "<br/>";
    echo $e->getFile() . "<br/>";
    echo $e->getLine() . "<br/>";
}

function fakeCpfGeneration(): string
{

    return rand(100, 999) . "." . rand(100, 999) . "." . rand(100, 900) . "-" . rand(10, 99);
}