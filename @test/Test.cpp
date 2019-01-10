//
// Created by Troldal on 2019-01-10.
//

#include "gtest/gtest.h"
#include <vector>
int main(int argc, char **argv) {
    ::testing::InitGoogleTest(&argc, argv);
    return RUN_ALL_TESTS();
}
TEST(example, sum_zero) {

    ASSERT_EQ(0, 0);
}
TEST(example, sum_five) {

    ASSERT_EQ(14, 15);
}
