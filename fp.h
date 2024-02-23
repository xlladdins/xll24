// fp.h - FP12 data type
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
extern "C" {
#include "fpx.h"
}
#include <algorithm>
#include <initializer_list>
//#include <mdspan>
#include <span>
#include <stdexcept>

constexpr int rows(const FP12& a) noexcept
{
	return a.rows;
}
constexpr int columns(const FP12& a) noexcept
{
	return a.columns;
}
constexpr int size(const FP12& a) noexcept
{
	return a.rows * a.columns;
}

constexpr double* array(FP12& a) noexcept
{
	return a.array;
}
constexpr const double* array(const FP12& a) noexcept
{
	return a.array;
}

constexpr auto span(FP12& a) noexcept
{
	return std::span<double>(array(a), size(a));
}
constexpr auto span(const FP12& a) noexcept
{
	return std::span<const double>(array(a), size(a));
}

constexpr auto row(FP12& a, int i) noexcept
{
	return std::span<double>(array(a) + i*columns(a), columns(a));
}
constexpr auto row(const FP12& a, int i) noexcept
{
	return std::span<const double>(array(a) + i * columns(a), columns(a));
}
/*
constexpr auto column(FP12& a, int j) noexcept
{
	return std::mdspan<double>(array(a) + j, rows(a));
}
constexpr auto column(const FP12& a, int j) noexcept
{
	return std::mdspan<const double>(array(a) + j, rows(a));
}
*/
constexpr double* begin(FP12& a) noexcept
{
	return array(a);
}
constexpr const double* begin(const FP12& a) noexcept
{
	return array(a);
}
constexpr double* end(FP12& a) noexcept
{
	return array(a) + size(a);
}
constexpr const double* end(const FP12& a) noexcept
{
	return array(a) + size(a);
}

namespace xll {
	class FPX {
		struct fpx* fpx_;
	public:
		FPX() noexcept
			: fpx_(nullptr)
		{ }
		FPX(int r, int c)
			: fpx_(fpx_malloc(r, c))
		{
			if (!fpx_) {
				throw std::runtime_error(__FUNCTION__ ": fpx_malloc failed");
			}
		}
		// Copy from array having at least r*c elements.
		FPX(int r, int c, const double* a)
			: FPX(r, c)
		{
			std::copy_n(a, r*c, fpx_->array);
		}
		FPX(std::initializer_list<double> a)
			: FPX(1, static_cast<int>(a.size()), a.begin())
		{ }
		FPX(const FPX& a)
			: FPX(rows(a), columns(a), array(a))
		{ }
		FPX(FPX&& a) noexcept
			: fpx_(a.fpx_)
		{
			a.fpx_ = nullptr;
		}
		FPX& operator=(const FPX& a)
		{
			if (this != &a) {
				fpx_free(fpx_);
				fpx_ = fpx_malloc(rows(a), columns(a));
				std::copy_n(array(a), size(a), fpx_->array);
			}

			return *this;
		}
		FPX& operator=(FPX&& a) noexcept
		{
			if (this != &a) {
				fpx_free(fpx_);
				fpx_ = std::exchange(a.fpx_, nullptr);
			}

			return *this;
		}
		~FPX()
		{
			fpx_free(fpx_);
		}

		explicit operator bool() const noexcept
		{
			return fpx_ != nullptr;
		}

		// Return pointer to _FP12.
		_FP12* get() noexcept
		{
			return reinterpret_cast<_FP12*>(fpx_);
		}
		const _FP12* get() const noexcept
		{
			return reinterpret_cast<const _FP12*>(fpx_);
		}
		operator _FP12& () noexcept
		{
			return reinterpret_cast<_FP12&>(*fpx_);
		}
		operator const _FP12&() const noexcept
		{
			return reinterpret_cast<const _FP12&>(*fpx_);
		}

		FPX& resize(int r, int c)
		{
			auto _fpx = fpx_realloc(fpx_, r, c);
			if (_fpx) {
				fpx_ = _fpx;
			}
			else {
				throw std::runtime_error(__FUNCTION__ ": fpx_realloc failed");
			}

			return *this;
		}
		constexpr FPX& reshape(int r, int c)
		{
			if (r * c != size(*this)) {
				throw std::runtime_error(__FUNCTION__ ": reshape must have same size");
			}
			
			fpx_->rows = r;
			fpx_->columns = c;

			return *this;
		}

		double operator[](int i) const noexcept
		{
			return fpx_->array[i];
		}
		double& operator[](int i) noexcept
		{
			return fpx_->array[i];
		}

		double operator()(int i, int j) const noexcept
		{
			return fpx_->array[fpx_index(fpx_, i, j)];
		}
		double& operator()(int i, int j) noexcept
		{
			return fpx_->array[fpx_index(fpx_, i, j)];
		}
	};
}

constexpr bool operator==(const _FP12& a, const _FP12& b)
{
	return a.rows == b.rows 
		&& a.columns == b.columns 
		&& std::equal(begin(a), end(a), begin(b), end(b));
}