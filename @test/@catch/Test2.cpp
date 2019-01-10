//
// Created by Troldal on 2019-01-10.
//

// this tells catch to provide a main()
// only do this in one cpp file

#define CATCH_CONFIG_MAIN
#include "catch.hpp"
#include <vector>
TEST_CASE("Sum of integers for a short vector", "[short]") {

    REQUIRE(15 == 15);
}

TEST_CASE("Sum of integers for a longer vector", "[long]") {

    REQUIRE(0 == 500500);
}
