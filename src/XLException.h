//
// Created by Kenneth Balslev on 06/12/2017.
//

#ifndef OPENXLEXE_XLEXCEPTION_H
#define OPENXLEXE_XLEXCEPTION_H

#include <stdexcept>

namespace RapidXLSX
{

    class XLException: std::runtime_error
    {
    public:
        explicit XLException(const std::string &err)
            : runtime_error(err)
        {}

    };

}


#endif //OPENXLEXE_XLEXCEPTION_H
