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

namespace xll {

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
		return std::span<double>(array(a) + i * columns(a), columns(a));
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

	class FPX {
		struct fpx* fpx_;
	public:
		FPX(int r = 0, int c = 0)
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
			std::copy_n(a, r * c, fpx_->array);
		}
		FPX(std::initializer_list<double> a)
			: FPX(1, static_cast<int>(a.size()), a.begin())
		{ }
		FPX(const FPX& a)
			: FPX(a.rows(), a.columns(), a.array())
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
				fpx_ = fpx_malloc(a.rows(), a.columns());
				std::copy_n(a.array(), a.size(), fpx_->array);
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

		int rows() const noexcept
		{
			return fpx_ ? fpx_->rows : 0;
		}
		int columns() const noexcept
		{
			return fpx_ ? fpx_->columns : 0;
		}
		int size() const noexcept
		{
			return rows() * columns();
		}
		double* array() noexcept
		{
			return fpx_ ? fpx_->array : nullptr;
		}
		const double* array() const noexcept
		{
			return fpx_ ? fpx_->array : nullptr;
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
		operator const _FP12& () const noexcept
		{
			return reinterpret_cast<const _FP12&>(*fpx_);
		}

		FPX& resize(int r, int c)
		{
			if (r * c != size()) {
				auto _fpx = fpx_realloc(fpx_, r, c);
				if (_fpx) {
					fpx_ = _fpx;
				}
				else {
					throw std::runtime_error(__FUNCTION__ ": fpx_realloc failed");
				}
			}
			else {
				fpx_->rows = r;
				fpx_->columns = c;
			}

			return *this;
		}

		FPX& push_back(const double* a, int n = 1)
		{
			if (a == nullptr) {
				throw std::runtime_error(__FUNCTION__ ": a is nullptr");
			}
			if (n < 1) {
				throw std::runtime_error(__FUNCTION__ ": n < 1");
			}
			int r = rows(), r_ = r;
			int c = columns(), c_ = c;
			if (c == 0) {
				r_ = n;
				c_ = 1;
			}
			else {
				r_ = r + 1 + (n - 1) / c;
			}
			auto _fpx = fpx_realloc(fpx_, r_, c ? c : 1);
			if (_fpx) {
				fpx_ = _fpx;
				std::copy_n(a, n, fpx_->array + r * c);
			}
			else {
				throw std::runtime_error(__FUNCTION__ ": fpx_realloc failed");
			}

			return *this;
		}
		FPX& push_back(double x)
		{
			return push_back(&x);
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

	// Fixed size array.
	template<size_t N, size_t M>
	struct fp12 {
		int rows = static_cast<int>(N);
		int columns = static_cast<int>(M);
		double array[N * M];
		fp12& reshape(int r, int c)
		{
			if (r * c != rows * columns) {
				rows = 0;
				columns = 0;
				array[0] = std::numeric_limits<double>::quiet_NaN();
			}
			rows = r;
			columns = c;

			return *this;
		}	
		// For use with stand-alone functions.
		_FP12* get() noexcept
		{
			return reinterpret_cast<_FP12*>(this);
		}
		const _FP12* get() const noexcept
		{
			return reinterpret_cast<const _FP12*>(this);
		}
	};
}

constexpr bool operator==(const _FP12& a, const _FP12& b)
{
	return a.rows == b.rows
		&& a.columns == b.columns
		&& std::equal(xll::begin(a), xll::end(a), xll::begin(b), xll::end(b));
}