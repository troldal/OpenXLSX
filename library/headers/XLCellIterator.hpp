////
//// Created by Kenneth Balslev on 05/07/2020.
////
//
//#ifndef OPENXLSX_XLCELLITERATOR_IMPL_HPP
//#define OPENXLSX_XLCELLITERATOR_IMPL_HPP
//
//#include "XLCellReference_Impl.hpp"
//#include "XLCell_Impl.hpp"
//#include "XLXml_Impl.hpp"
//
// namespace OpenXLSX
//{
//    enum class XLIteratorDirection { Forward, Reverse };
//
//    class XLCellIterator
//    {
//    public:
//        using iterator_category = std::forward_iterator_tag;
//        using value_type = XLCell;
//        using difference_type = int64_t;
//        using pointer = XLCell*;
//        using reference = XLCell&;
//
//        XLCellIterator(XLCellReference cellRef, XMLNode dataNode) {}
//        XLCellIterator& operator++() { self_type i = *this; ptr_++; return i; }
//        XLCellIterator operator++(int junk) { ptr_++; return *this; }
//        reference operator*() { return m_currentCell; }
//        pointer operator->() { return &m_currentCell; }
//        bool operator==(const self_type& rhs) { return ptr_ == rhs.ptr_; }
//        bool operator!=(const self_type& rhs) { return ptr_ != rhs.ptr_; }
//
//
//    private:
//
//        XLCellReference   m_cellReference;
//        XMLNode           m_currentRowNode;
//        XMLNode           m_currentCellNode;
//        XMLNode             m_previousCellNode;
//        XLCell              m_currentCell;
//        XLIteratorDirection m_iteratorDirection;
//    };
//
//}  // namespace OpenXLSX::Impl
//
//#endif //OPENXLSX_XLCELLITERATOR_IMPL_HPP
