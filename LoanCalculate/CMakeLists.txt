cmake_minimum_required(VERSION 3.9)
project(LoanCalculate)
if(WIN32)
    set(CMAKE_CXX_FLAGS "-static-libgcc -static-libstdc++")
endif(WIN32)
set(CMAKE_CXX_STANDARD 11)

add_executable(LoanCalculate main.cpp DataStructure.h DataStructure.cpp)