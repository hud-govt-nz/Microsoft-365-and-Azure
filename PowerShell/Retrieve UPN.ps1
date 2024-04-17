$Email =@()
$Names = @(
"Saera Chun",
"Joseph Winton",
"Nitika Sehgal",
"Kyle Worsley",
"Brandon Wise",
"Pang Khun",
"Simon Morris",
"Ruth Holness",
"Carl Robinson",
"Paul Johnstone",
"Roisin Lamar",
"Jessie Wilson",
"Emma Funnell",
"Miranda Devlin",
"Biddy Livesey",
"Ariel McLean-Robinson",
"Emily Shrosbree",
"Zohreh Karaminejad",
"Rosalind Dibley",
"Keriata Stuart",
"Nathan Greenough",
"Katie Wittkowski",
"Alex Gunn",
"Adam Kado",
"Keith Ng",
"Katrina Buxton",
"Helen Cox",
"Eimear Doyle",
"Richard Deakin",
"Mariona Roige Valiente",
"Craig Fredrickson",
"Katie Wellington",
"Brittany Goodwin",
"Kate Reid"


)





foreach ($Name in $Names) {
    $User = Get-MgUser -Filter "displayName eq '$Name'"
    $DisplayName = $User.DisplayName
    $UPN = $User.UserPrincipalName

    $Email += [PSCustomObject]@{
        #DisplayName = $DisplayName
        UserPrincipalName = $UPN
    }
}

Foreach ($i in $Email) {
    $I.UserPrincipalName
}