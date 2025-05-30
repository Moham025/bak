#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Conversion de nombres en lettres - Portage du code VBA vers Python
Devise=0   aucune
      =1   Euro €
      =2   Dollar $
Langue=0   Français
      =1   Belgique
      =2   Suisse

Conversion limitée à 999 999 999 999 999 ou 9 999 999 999 999,99
Si le nombre contient plus de 2 décimales, il est arrondi à 2 décimales
"""

import math


def conv_number_letter(nombre, devise=0, langue=0):
    """
    Fonction principale de conversion de nombre en lettres
    
    Args:
        nombre (float): Le nombre à convertir
        devise (int): 0=aucune, 1=Euro, 2=Dollar
        langue (int): 0=Français, 1=Belgique, 2=Suisse
    
    Returns:
        str: Le nombre en lettres
    """
    try:
        nombre = float(nombre)
    except:
        return "#ErreurConversion"
    
    b_negatif = False
    if nombre < 0:
        b_negatif = True
        nombre = abs(nombre)
    
    dbl_ent = int(nombre)
    by_dec = int(round((nombre - dbl_ent) * 100))
    
    # Vérification des limites
    if by_dec == 0:
        if dbl_ent > 999999999999999:
            return "#TropGrand"
    else:
        if dbl_ent > 9999999999999.99:
            return "#TropGrand"
    
    # Configuration des devises
    str_dev = ""
    str_centimes = ""
    
    if devise == 0:
        if by_dec > 0:
            str_dev = " virgule"
    elif devise == 1:
        str_dev = f" ( {nombre} )" + " Francs CFA"
        if by_dec > 0:
            str_centimes = " Cents"
    elif devise == 2:
        str_dev = " Dollar"
        if by_dec > 0:
            str_centimes = " Cent"
    
    if dbl_ent > 1 and devise != 0:
        str_dev = str_dev  # Pas de modification pour les pluriels dans cette version
    
    result = conv_num_ent(float(dbl_ent), langue) + str_dev
    if by_dec > 0:
        result += " " + conv_num_dizaine(by_dec, langue) + str_centimes
    
    if b_negatif:
        result = "moins " + result
    
    return result.strip()


def conv_num_ent(nombre, langue):
    """
    Conversion des entiers (gestion des milliers, millions, etc.)
    """
    if nombre == 0:
        return "zéro"
    
    result = ""
    
    # Unités
    i_tmp = int(nombre % 1000)
    result = conv_num_cent(i_tmp, langue)
    
    # Milliers
    nombre = int(nombre / 1000)
    i_tmp = int(nombre % 1000)
    if i_tmp > 0:
        str_tmp = conv_num_cent(i_tmp, langue)
        if i_tmp == 1:
            str_tmp = "mille "
        else:
            str_tmp = str_tmp + " mille "
        result = str_tmp + result
    
    # Millions
    nombre = int(nombre / 1000)
    i_tmp = int(nombre % 1000)
    if i_tmp > 0:
        str_tmp = conv_num_cent(i_tmp, langue)
        if i_tmp == 1:
            str_tmp = str_tmp + " million "
        else:
            str_tmp = str_tmp + " millions "
        result = str_tmp + result
    
    # Milliards
    nombre = int(nombre / 1000)
    i_tmp = int(nombre % 1000)
    if i_tmp > 0:
        str_tmp = conv_num_cent(i_tmp, langue)
        if i_tmp == 1:
            str_tmp = str_tmp + " milliard "
        else:
            str_tmp = str_tmp + " milliards "
        result = str_tmp + result
    
    # Billions
    nombre = int(nombre / 1000)
    i_tmp = int(nombre % 1000)
    if i_tmp > 0:
        str_tmp = conv_num_cent(i_tmp, langue)
        if i_tmp == 1:
            str_tmp = str_tmp + " billion "
        else:
            str_tmp = str_tmp + " billions "
        result = str_tmp + result
    
    return result.strip()


def conv_num_dizaine(nombre, langue):
    """
    Conversion des dizaines (0-99)
    """
    if not (0 <= nombre <= 99):
        return ""
    
    tab_unit = ["", "un", "deux", "trois", "quatre", "cinq", "six", "sept",
                "huit", "neuf", "dix", "onze", "douze", "treize", "quatorze", 
                "quinze", "seize", "dix-sept", "dix-huit", "dix-neuf"]
    
    tab_diz = ["", "", "vingt", "trente", "quarante", "cinquante",
               "soixante", "soixante", "quatre-vingt", "quatre-vingt"]
    
    # Adaptations selon la langue
    if langue == 1:  # Belgique
        tab_diz[7] = "septante"
        tab_diz[9] = "nonante"
    elif langue == 2:  # Suisse
        tab_diz[7] = "septante"
        tab_diz[8] = "huitante"
        tab_diz[9] = "nonante"
    
    by_diz = int(nombre / 10)
    by_unit = nombre % 10
    str_liaison = "-"
    
    if by_unit == 1:
        str_liaison = " "
    
    # Cas spéciaux
    if by_diz == 0:
        str_liaison = ""
    elif by_diz == 1:
        by_unit = by_unit + 10
        str_liaison = ""
    elif by_diz == 7:
        if langue == 0:  # Français
            by_unit = by_unit + 10
    elif by_diz == 8:
        if langue != 2:  # Pas suisse
            str_liaison = "-"
    elif by_diz == 9:
        if langue == 0:  # Français
            by_unit = by_unit + 10
            str_liaison = "-"
    
    result = tab_diz[by_diz]
    
    # Cas spécial pour quatre-vingt
    if by_diz == 8 and langue != 2 and by_unit == 0:
        result = result + "s"
    
    # Ajout de l'unité
    if by_unit < len(tab_unit) and tab_unit[by_unit] != "":
        if by_unit == 1 and by_diz in [2, 3, 4, 5, 6] and str_liaison == " ":
            result = result + " et " + tab_unit[by_unit]
        else:
            result = result + str_liaison + tab_unit[by_unit]
    
    return result


def conv_num_cent(nombre, langue):
    """
    Conversion des centaines (0-999)
    """
    if not (0 <= nombre <= 999):
        return ""
    
    tab_unit = ["", "un", "deux", "trois", "quatre", "cinq", "six", "sept",
                "huit", "neuf", "dix"]
    
    by_cent = int(nombre / 100)
    by_reste = nombre % 100
    str_reste = conv_num_dizaine(by_reste, langue)
    
    if by_cent == 0:
        result = str_reste
    elif by_cent == 1:
        if by_reste == 0:
            result = "cent"
        else:
            result = "cent " + str_reste
    else:
        if by_reste == 0:
            result = tab_unit[by_cent] + " cents"
        else:
            result = tab_unit[by_cent] + " cent " + str_reste
    
    return result


# Tests si le fichier est exécuté directement
if __name__ == "__main__":
    # Tests
    test_cases = [
        (123, 0, 0),
        (1750000, 1, 0),
        (123.45, 2, 0),
        (80, 0, 0),
        (81, 0, 1),
        (91, 0, 2)
    ]
    
    for nombre, devise, langue in test_cases:
        result = conv_number_letter(nombre, devise, langue)
        print(f"{nombre} (devise={devise}, langue={langue}) -> {result}") 