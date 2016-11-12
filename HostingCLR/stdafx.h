// stdafx.h : fichier Include pour les fichiers Include système standard,
// ou les fichiers Include spécifiques aux projets qui sont utilisés fréquemment,
// et sont rarement modifiés
//

#pragma once

#include "targetver.h"

#include <stdio.h>
#include <tchar.h>



// TODO: faites référence ici aux en-têtes supplémentaires nécessaires au programme
#include <metahost.h>
#pragma comment(lib, "MSCorEE.lib")

#import "C:\Windows\Microsoft.NET\Framework\v2.0.50727\mscorlib.tlb" raw_interfaces_only \
    high_property_prefixes("_get","_put","_putref")		\
    rename("ReportEvent", "InteropServices_ReportEvent")
using namespace mscorlib;