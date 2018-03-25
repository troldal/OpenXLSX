//
// Created by KBA012 on 20-12-2017.
//

#ifndef OPENXLEXE_XLCELLFORMULA_H
#define OPENXLEXE_XLCELLFORMULA_H

#include <XML/XMLNode.h>

namespace RapidXLSX
{

    class XLCellFormula
    {
    public:


    private:
        XMLNode *m_formulaNode;
        std::string m_formula;

    };

}


#endif //OPENXLEXE_XLCELLFORMULA_H
