//
// Created by kenne on 24/05/2024.
//

#pragma once

#include <memory>

namespace cpp
{
    template<typename T, typename... Args>
    std::unique_ptr<T> make_unique(Args&&... args)
    {
        return std::unique_ptr<T>(new T(std::forward<Args>(args)...));
    }
}