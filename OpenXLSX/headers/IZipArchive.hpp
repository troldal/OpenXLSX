/*

   ____                               ____      ___ ____       ____  ____      ___
  6MMMMb                              `MM(      )M' `MM'      6MMMMb\`MM(      )M'
 8P    Y8                              `MM.     d'   MM      6M'    ` `MM.     d'
6M      Mb __ ____     ____  ___  __    `MM.   d'    MM      MM        `MM.   d'
MM      MM `M6MMMMb   6MMMMb `MM 6MMb    `MM. d'     MM      YM.        `MM. d'
MM      MM  MM'  `Mb 6M'  `Mb MMM9 `Mb    `MMd       MM       YMMMMb     `MMd
MM      MM  MM    MM MM    MM MM'   MM     dMM.      MM           `Mb     dMM.
MM      MM  MM    MM MMMMMMMM MM    MM    d'`MM.     MM            MM    d'`MM.
YM      M9  MM    MM MM       MM    MM   d'  `MM.    MM            MM   d'  `MM.
 8b    d8   MM.  ,M9 YM    d9 MM    MM  d'    `MM.   MM    / L    ,M9  d'    `MM.
  YMMMM9    MMYMMM9   YMMMM9 _MM_  _MM_M(_    _)MM_ _MMMMMMM MYMMMM9 _M(_    _)MM_
            MM
            MM
           _MM_

  Copyright (c) 2022, Kenneth Troldal Balslev

  All rights reserved.

  Redistribution and use in source and binary forms, with or without
  modification, are permitted provided that the following conditions are met:
  - Redistributions of source code must retain the above copyright
    notice, this list of conditions and the following disclaimer.
  - Redistributions in binary form must reproduce the above copyright
    notice, this list of conditions and the following disclaimer in the
    documentation and/or other materials provided with the distribution.
  - Neither the name of the author nor the
    names of any contributors may be used to endorse or promote products
    derived from this software without specific prior written permission.

  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
  ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
  DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
  (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
  LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
  (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
  SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

 */

#ifndef OPENXLSX_IZIPARCHIVE_HPP
#define OPENXLSX_IZIPARCHIVE_HPP

#pragma warning(push)
#pragma warning(disable : 4251)
#pragma warning(disable : 4275)

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"

#include <memory>
#include <string>

namespace OpenXLSX
{
    /**
     * @brief This class functions as a wrapper around any class that provides the necessary functionality for
     * a zip archive.
     * @details This class works by applying 'type erasure'. This enables the use of objects of any class, the only
     * requirement being that it provides the right interface. No inheritance from a base class is needed.
     */
    class OPENXLSX_EXPORT IZipArchive
    {
    public:
        /**
         * @brief Default constructor
         */
        IZipArchive() : m_zipArchive() {} // NOLINT

        /**
         * @brief Constructor, taking the target object as an argument.
         * @tparam T The type of the target object (will be auto deducted)
         * @param x The target object
         * @note This method is deliberately not marked 'explicit', because as a templated constructor, it should be able
         * to take any type as an argument. However, only objects that satisfy the required interface can be used.
         */
        template<typename T>
        IZipArchive(const T& zipArchive) : m_zipArchive { std::make_unique<Model<T>>(zipArchive) } {} // NOLINT

        /**
         * @brief Copy constructor
         * @param other
         */
        IZipArchive(const IZipArchive& other) : m_zipArchive(other.m_zipArchive ? other.m_zipArchive->clone() : nullptr) {}

        /**
         * @brief Move constructor
         * @param other
         */
        IZipArchive(IZipArchive&& other) noexcept = default;

        /**
         * @brief Destructor
         */
        ~IZipArchive() = default;

        /**
         * @brief
         * @tparam T
         * @param x
         * @return
         */
        template<typename T>
        inline IZipArchive& operator=(const T& zipArchive)
        {
            m_zipArchive = std::make_unique<Model<T>>(zipArchive);
            return *this;
        }

        /**
         * @brief
         * @param other
         * @return
         */
        inline IZipArchive& operator=(const IZipArchive& other)
        {
            IZipArchive copy(other);
            *this = std::move(copy);
            return *this;
        }

        /**
         * @brief
         * @param other
         * @return
         */
        inline IZipArchive& operator=(IZipArchive&& other) noexcept = default;

        /**
         * @brief
         * @return
         */
        inline explicit operator bool() const
        {
            return isValid();
        }

        inline bool isValid() const {
            return m_zipArchive->isValid();

        }

        inline bool isOpen() const {
            return m_zipArchive->isOpen();
        }

        inline void open(const std::string& fileName) {
            m_zipArchive->open(fileName);
        }

        inline void close() const {
            m_zipArchive->close();
        }

        inline void save(const std::string& path) {
            m_zipArchive->save(path);
        }

        inline void addEntry(const std::string& name, const std::string& data) {
            m_zipArchive->addEntry(name, data);
        }

        inline void deleteEntry(const std::string& entryName) {
            m_zipArchive->deleteEntry(entryName);
        }

        inline std::string getEntry(const std::string& name) {
            return m_zipArchive->getEntry(name);
        }

        inline bool hasEntry(const std::string& entryName) {
            return m_zipArchive->hasEntry(entryName);
        }

    private:
        /**
         * @brief
         */
        struct Concept
        {
        public:
            /**
             * @brief
             */
            Concept() = default;

            /**
             * @brief
             */
            Concept(const Concept&) = default;

            /**
             * @brief
             */
            Concept(Concept&&) noexcept = default;

            /**
             * @brief
             */
            virtual ~Concept() = default;

            /**
             * @brief
             * @return
             */
            inline Concept& operator=(const Concept&) = default;

            /**
             * @brief
             * @return
             */
            inline Concept& operator=(Concept&&) noexcept = default;

            /**
             * @brief
             * @return
             */
            inline virtual std::unique_ptr<Concept> clone() const = 0;

            inline virtual bool isValid() const = 0;

            inline virtual bool isOpen() const = 0;

            inline virtual void open(const std::string& fileName) = 0;

            inline virtual void close() = 0;

            inline virtual void save (const std::string& path) = 0;

            inline virtual void addEntry(const std::string& name, const std::string& data) = 0;

            inline virtual void deleteEntry(const std::string& entryName) = 0;

            inline virtual std::string getEntry(const std::string& name) = 0;

            inline virtual bool hasEntry(const std::string& entryName) = 0;

        };

        /**
         * @brief
         * @tparam T
         */
        template<typename T>
        struct Model : Concept
        {
        public:
            /**
             * @brief
             * @param x
             */
            explicit Model(const T& x) : ZipType(x) {}

            /**
             * @brief
             * @param other
             */
            Model(const Model& other) = default;

            /**
             * @brief
             * @param other
             */
            Model(Model&& other) noexcept = default;

            /**
             * @brief
             */
            ~Model() override = default;

            /**
             * @brief
             * @param other
             * @return
             */
            inline Model& operator=(const Model& other) = default;

            /**
             * @brief
             * @param other
             * @return
             */
            inline Model& operator=(Model&& other) noexcept = default;

            /**
             * @brief
             * @return
             */
            inline std::unique_ptr<Concept> clone() const override
            {
                return std::make_unique<Model<T>>(ZipType);
            }

            inline bool isValid() const override {
                return ZipType.isValid();
            }

            inline bool isOpen() const override {
                return ZipType.isOpen();
            }

            inline void open(const std::string& fileName) override {
                ZipType.open(fileName);
            }

            inline void close() override {
                ZipType.close();
            }

            inline void save(const std::string& path) override {
                ZipType.save(path);
            }

            inline void addEntry(const std::string& name, const std::string& data) override {
                ZipType.addEntry(name, data);
            }

            inline void deleteEntry(const std::string& entryName) override {
                ZipType.deleteEntry(entryName);
            }

            inline std::string getEntry(const std::string& name) override {
                return ZipType.getEntry(name);
            }

            inline bool hasEntry(const std::string& entryName) override {
                return ZipType.hasEntry(entryName);
            }

        private:
            T ZipType;
        };

        std::unique_ptr<Concept> m_zipArchive;
    };

}    // namespace OpenXLSX

#pragma warning(pop)
#endif    // OPENXLSX_IZIPARCHIVE_HPP
