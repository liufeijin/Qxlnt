// Copyright (c) 2014-2020 Thomas Fussell
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, WRISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE
//
// @license: http://www.opensource.org/licenses/mit-license.php
// @author: see AUTHORS file

#include <xlnt/cell/cell.hpp>
#include <xlnt/styles/style.hpp>
#include <xlnt/workbook/workbook.hpp>
#include <xlnt/worksheet/range.hpp>
#include <xlnt/worksheet/range_iterator.hpp>
#include <xlnt/worksheet/range_reference.hpp>
#include <xlnt/worksheet/worksheet.hpp>
#include <xlnt/worksheet/column_properties.hpp>
#include <xlnt/worksheet/row_properties.hpp>
#include <xlnt/styles/fill.hpp>


namespace xlnt {

range::range(class worksheet ws, const range_reference &reference, major_order order, bool skip_null)
    : ws_(ws),
      ref_(reference),
      order_(order),
      skip_null_(skip_null)
{
}

range::~range() = default;

void range::clear_cells()
{
    if (ref_.top_left().column() == ws_.lowest_column()
        && ref_.bottom_right().column() == ws_.highest_column())
    {
        for (auto row = ref_.top_left().row(); row <= ref_.bottom_right().row(); ++row)
        {
            ws_.clear_row(row);
        }
    }
    else
    {
        for (auto row = ref_.top_left().row(); row <= ref_.bottom_right().row(); ++row)
        {
            for (auto column = ref_.top_left().column(); column <= ref_.bottom_right().column(); ++column)
            {
                ws_.clear_cell(xlnt::cell_reference(column, row));
            }
        }
    }
}

cell_vector range::operator[](std::size_t index)
{
    return vector(index);
}

const cell_vector range::operator[](std::size_t index) const
{
    return vector(index);
}

const worksheet &range::target_worksheet() const
{
    return ws_;
}

range_reference range::reference() const
{
    return ref_;
}

std::size_t range::length() const
{
    if (order_ == major_order::row)
    {
        return ref_.bottom_right().row() - ref_.top_left().row() + 1;
    }

    return (ref_.bottom_right().column() - ref_.top_left().column()).index + 1;
}

bool range::operator==(const range &comparand) const
{
    return ref_ == comparand.ref_
        && ws_ == comparand.ws_
        && order_ == comparand.order_;
}

cell_vector range::vector(std::size_t vector_index)
{
    auto cursor = ref_.top_left();

    if (order_ == major_order::row)
    {
        cursor.row(cursor.row() + static_cast<row_t>(vector_index));
    }
    else
    {
        cursor.column_index(cursor.column_index() + static_cast<column_t::index_t>(vector_index));
    }

    return cell_vector(ws_, cursor, ref_, order_, skip_null_, false);
}

const cell_vector range::vector(std::size_t vector_index) const
{
    auto cursor = ref_.top_left();

    if (order_ == major_order::row)
    {
        cursor.row(cursor.row() + static_cast<row_t>(vector_index));
    }
    else
    {
        cursor.column_index(cursor.column_index() + static_cast<column_t::index_t>(vector_index));
    }

    return cell_vector(ws_, cursor, ref_, order_, skip_null_, false);
}

const cell_vector range::vector(const xlnt::horizontal_alignment dir) const
{
    auto cursor = ref_.top_left();
    if (dir == xlnt::horizontal_alignment::left) {     
        cursor.column_index(cursor.column_index());
    } else if (dir == xlnt::horizontal_alignment::right) {     
        cursor.column_index(cursor.column_index() + (ref_.bottom_right().column() - ref_.top_left().column()).index);
    }
    
    return cell_vector(ws_, cursor, ref_, major_order::column, skip_null_, false);
}
const cell_vector range::vector(const xlnt::vertical_alignment dir) const
{
    auto cursor = ref_.top_left();
    if (dir == xlnt::vertical_alignment::top) {
        ;
    } else if (dir == xlnt::vertical_alignment::bottom) {
        cursor.row(cursor.row() + (ref_.bottom_right().row() - ref_.top_left().row()));
    }
    
    return cell_vector(ws_, cursor, ref_, major_order::row, skip_null_, false);
}

bool range::contains(const cell_reference &ref)
{
    return ref_.top_left().column_index() <= ref.column_index()
        && ref_.bottom_right().column_index() >= ref.column_index()
        && ref_.top_left().row() <= ref.row()
        && ref_.bottom_right().row() >= ref.row();
}

range range::alignment(const xlnt::alignment &new_alignment)
{
    apply([&new_alignment](class cell c) { c.alignment(new_alignment); });
    return *this;
}

range range::border(const xlnt::border &new_border)
{
    apply([&new_border](class cell c) { c.border(new_border); });
    return *this;
}

void range::border_style(const xlnt::horizontal_alignment dir, xlnt::border_style bs)
{
    auto cells = vector(dir);
    for(auto it=cells.begin(); it!=cells.end(); ++it) {
        auto curBorder = (*it).border();

        if (dir == horizontal_alignment::left) {
            optional<border::border_property> obp = curBorder.side(border_side::start);
            border::border_property bp = obp.get();
            bp.style().set(bs);
            curBorder.side(border_side::start, bp);
        } else if (dir == horizontal_alignment::right) {
            optional<border::border_property> obp = curBorder.side(border_side::end);
            border::border_property bp = obp.get();
            bp.style().set(bs);
            curBorder.side(border_side::end, bp);
        }

        (*it).border(curBorder);
    }
}
void range::border_style(const xlnt::vertical_alignment dir, xlnt::border_style bs)
{
    auto cells = vector(dir);
    for(auto it=cells.begin(); it!=cells.end(); ++it) {
        auto curBorder = (*it).border();

        if (dir == vertical_alignment::top) {
            optional<border::border_property> obp = curBorder.side(border_side::top);
            border::border_property bp = obp.get();
            bp.style().set(bs);
            curBorder.side(border_side::top, bp);
        } else if (dir == vertical_alignment::bottom) {
            optional<border::border_property> obp = curBorder.side(border_side::bottom);
            border::border_property bp = obp.get();
            bp.style().set(bs);
            curBorder.side(border_side::bottom, bp);
        }

        (*it).border(curBorder);
    }
}

range range::fill(const xlnt::fill &new_fill)
{
    apply([&new_fill](class cell c) { c.fill(new_fill); });
    return *this;
}

range range::font(const xlnt::font &new_font)
{
    apply([&new_font](class cell c) { c.font(new_font); });
    return *this;
}
range& range::font_size(const double new_font_size)
{
    apply([new_font_size](class cell c) { 
        auto old_font = c.font(); 
        old_font.size(new_font_size);
        c.font(old_font); 
    });

    return *this;
}
range& range::color(const double color)
{
    apply([color](class cell c) { 
        xlnt::style s = c.style();
        xlnt::pattern_fill pf;
        pf.type(xlnt::pattern_fill_type::none);
        pf.background(xlnt::color::white());        
        s.fill(xlnt::fill(pf));
        c.style(s); 
    });

    return *this;
}

range& range::clear_value()
{
    apply([](class cell c) { c.clear_value(); });
    return *this;
}


range range::number_format(const xlnt::number_format &new_number_format)
{
    apply([&new_number_format](class cell c) { c.number_format(new_number_format); });
    return *this;
}

range range::protection(const xlnt::protection &new_protection)
{
    apply([&new_protection](class cell c) { c.protection(new_protection); });
    return *this;
}

range range::style(const class style &new_style)
{
    apply([&new_style](class cell c) { c.style(new_style); });
    return *this;
}

range range::style(const std::string &style_name)
{
    return style(ws_.workbook().style(style_name));
}

conditional_format range::conditional_format(const condition &when)
{
    return ws_.conditional_format(ref_, when);
}

void range::apply(std::function<void(class cell)> f)
{
    for (auto row : *this)
    {
        for (auto cell : row)
        {
            f(cell);
        }
    }
}

cell range::cell(const cell_reference &ref)
{
    return (*this)[ref.row() - 1][ref.column().index - 1];
}

const cell range::cell(const cell_reference &ref) const
{
    return (*this)[ref.row() - 1][ref.column().index - 1];
}

cell_vector range::front()
{
    return *begin();
}

const cell_vector range::front() const
{
    return *cbegin();
}

cell_vector range::back()
{
    return *(--end());
}

const cell_vector range::back() const
{
    return *(--cend());
}

range::iterator range::begin()
{
    return iterator(ws_, ref_.top_left(), ref_, order_, skip_null_);
}

range::iterator range::end()
{
    auto cursor = ref_.top_left();

    if (order_ == major_order::row)
    {
        cursor.row(ref_.bottom_right().row() + 1);
    }
    else
    {
        cursor.column_index(ref_.bottom_right().column_index() + 1);
    }

    return iterator(ws_, cursor, ref_, order_, skip_null_);
}

range::const_iterator range::cbegin() const
{
    return const_iterator(ws_, ref_.top_left(), ref_, order_, skip_null_);
}

range::const_iterator range::cend() const
{
    auto cursor = ref_.top_left();

    if (order_ == major_order::row)
    {
        cursor.row(ref_.bottom_right().row() + 1);
    }
    else
    {
        cursor.column_index(ref_.bottom_right().column_index() + 1);
    }

    return const_iterator(ws_, cursor, ref_, order_, skip_null_);
}

bool range::operator!=(const range &comparand) const
{
    return !(*this == comparand);
}

range::const_iterator range::begin() const
{
    return cbegin();
}

range::const_iterator range::end() const
{
    return cend();
}

range::reverse_iterator range::rbegin()
{
    return reverse_iterator(end());
}

range::reverse_iterator range::rend()
{
    return reverse_iterator(begin());
}

range::const_reverse_iterator range::crbegin() const
{
    return const_reverse_iterator(cend());
}

range::const_reverse_iterator range::rbegin() const
{
    return crbegin();
}

range::const_reverse_iterator range::crend() const
{
    return const_reverse_iterator(cbegin());
}

range::const_reverse_iterator range::rend() const
{
    return crend();
}


void range::column_width(double width)
{
	auto curOrder = order();
	order(xlnt::major_order::column);

    for(std::size_t i=0; i<length(); i++) {
        if (vector(i).empty()) {
            break;
        }

		auto ci = vector(i).front().reference().column_index();
        column_properties& col_prop = ws_.column_properties(ci);
		col_prop.width = width;
        //ws_.add_column_properties(ci, col_prop);
	}	

	order(curOrder);
}
void range::row_height(double h)
{
	auto curOrder = order();
	order(xlnt::major_order::row);

    for(std::size_t i=0; i<length(); i++) {
        if (vector(i).empty()) {
            break;
        }

		auto ci = vector(i).front().reference().row();
		row_properties& col_prop = ws_.row_properties(ci);
		col_prop.height = h;
		//ws_.add_row_properties(ci, col_prop);	
	}	

	order(curOrder);
}

} // namespace xlnt
