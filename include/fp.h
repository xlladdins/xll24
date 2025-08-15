// fp.h - FP12 data type
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <cstdint>
#include <algorithm>
#include <initializer_list>
//#include <mdspan>
#include <span>
#include <stdexcept>
#include "ensure.h"
extern "C" {
#include "fpx.h"
}

namespace xll {

	constexpr int32_t rows(const _FP12& a) noexcept
	{
		return a.rows;
	}
	constexpr int32_t columns(const _FP12& a) noexcept
	{
		return a.columns;
	}
	constexpr int32_t size(const _FP12& a) noexcept
	{
		return a.rows * a.columns;
	}

	constexpr double* array(_FP12& a) noexcept
	{
		return a.array;
	}
	constexpr const double* array(const _FP12& a) noexcept
	{
		return a.array;
	}

	constexpr double& index(_FP12& a, int32_t i, int32_t j) noexcept
	{
		return a.array[i * columns(a) + j];
	}
	constexpr double index(const _FP12& a, int32_t i, int32_t j) noexcept
	{
		return a.array[i * columns(a) + j];
	}

	constexpr auto span(_FP12& a) noexcept
	{
		return std::span<double>(array(a), size(a));
	}
	constexpr auto span(const _FP12& a) noexcept
	{
		return std::span<const double>(array(a), size(a));
	}

	constexpr auto row(_FP12& a, int32_t i) noexcept
	{
		return std::span<double>(array(a) + i * columns(a), columns(a));
	}
	constexpr auto row(const _FP12& a, int32_t i) noexcept
	{
		return std::span<const double>(array(a) + i * columns(a), columns(a));
	}
	/*
	constexpr auto column(_FP12& a, int32_t j) noexcept
	{
		return std::mdspan<double>(array(a) + j, rows(a));
	}
	constexpr auto column(const _FP12& a, int32_t j) noexcept
	{
		return std::mdspan<const double>(array(a) + j, rows(a));
	}
	*/
	constexpr double* begin(_FP12& a) noexcept
	{
		return array(a);
	}
	constexpr const double* begin(const _FP12& a) noexcept
	{
		return array(a);
	}
	constexpr double* end(_FP12& a) noexcept
	{
		return array(a) + size(a);
	}
	constexpr const double* end(const _FP12& a) noexcept
	{
		return array(a) + size(a);
	}
	// inplace transpose
	inline FP12& transpose(FP12& x) noexcept
	{
		fpx_transpose(reinterpret_cast<struct fpx*>(&x));

		return x;
	}

	class FPX {
		struct fpx* fpx_;
	public:
		FPX(int32_t r = 0, int32_t c = 0)
			: fpx_(fpx_malloc(r, c))
		{
			ensure(fpx_);
		}
		// Copy from array having at least r*c elements.
		FPX(int32_t r, int32_t c, const double* a)
			: FPX(r, c)
		{
			std::copy_n(a, r * c, fpx_->array);
		}
		FPX(const _FP12& a)
			: FPX(a.rows, a.columns, a.array)
		{ }
		// One row array.
		FPX(std::initializer_list<double> a)
			: FPX(1, static_cast<int32_t>(a.size()), a.begin())
		{ }
		template<size_t N>
		FPX(const double(&a)[N])
			: FPX(1, static_cast<int32_t>(N), a)
		{ }
		FPX(const FPX& a)
			: FPX(a.rows(), a.columns(), a.array())
		{ }
		// Construct from iterable
		template<class I>
			requires std::is_same_v<double, std::iter_value_t<I>>
		FPX(I i)
		{
			while (i) {
				append(*i);
				++i;
			}
		}
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
		FPX& operator=(const _FP12& a)
		{
			fpx_free(fpx_);
			fpx_ = fpx_malloc(a.rows, a.columns);
			std::copy_n(a.array, xll::size(a), fpx_->array);

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

		int32_t rows() const noexcept
		{
			return fpx_ ? fpx_->rows : 0;
		}
		int32_t columns() const noexcept
		{
			return fpx_ ? fpx_->columns : 0;
		}
		int32_t size() const noexcept
		{
			return rows() * columns();
		}
		double* array() noexcept
		{
			return (size() && fpx_) ? fpx_->array : nullptr;
		}
		const double* array() const noexcept
		{
			return (size() && fpx_) ? fpx_->array : nullptr;
		}

		// Return pointer to _FP12.
		_FP12* get() noexcept
		{
			return (size() && fpx_) ? reinterpret_cast<_FP12*>(fpx_) : nullptr;
		}
		const _FP12* get() const noexcept
		{
			return (size() && fpx_) ? reinterpret_cast<const _FP12*>(fpx_) : nullptr;
		}
		operator _FP12& () noexcept
		{
			return reinterpret_cast<_FP12&>(*fpx_);
		}
		operator const _FP12& () const noexcept
		{
			return reinterpret_cast<const _FP12&>(*fpx_);
		}

		double& operator[](int32_t i) noexcept
		{
			return array()[i];
		}
		double operator[](int32_t i) const noexcept
		{
			return array()[i];
		}
		double& operator()(int32_t i, int32_t j) noexcept
		{
			return xll::index(*get(), i, j);
		}
		double operator()(int32_t i, int32_t j) const noexcept
		{
			return xll::index(*get(), i, j);
		}

		FPX& resize(int32_t r, int32_t c)
		{
			if (r * c != size()) {
				auto _fpx = fpx_realloc(fpx_, r, c);
				ensure(_fpx);
				fpx_ = _fpx;
			}
			else {
				fpx_->rows = r;
				fpx_->columns = c;
			}

			return *this;
		}

		FPX& swap(FPX& a)
		{
			std::swap(fpx_, a.fpx_);

			return *this;
		}

		FPX& transpose()
		{
			xll::transpose(*get());

			return *this;
		}

		FPX& vstack(const _FP12& a)
		{
			if (size() == 0) {
				operator=(a);
			}
			else {
				ensure(columns() == a.columns);

				int32_t n = size();
				resize(rows() + a.rows, columns());
				std::copy_n(a.array, a.rows * a.columns, fpx_->array + n);
			}

			return *this;
		}
		FPX& vstack(const FPX& a)
		{
			return vstack(*a.get());
		}

		FPX& hstack(const _FP12& a)
		{
			if (size() == 0) {
				operator=(a);
			}
			else {
				// !!! not efficient
				FPX a_(*this);
				a_.transpose();
				FPX _a(a);
				_a.transpose();
				a_.vstack(_a);
				a_.transpose();
				swap(a_);
			}

			return *this;
		}
		FPX& hstack(const FPX& a)
		{
			return hstack(*a.get());
		}

		// Only works for vector arrays
		FPX& append(double x)
		{
			auto n = size();
			ensure(n == 0 || rows() == 1 || columns() == 1);

			if (n == 0) {
				resize(1, 1);
				operator[](0) = x;
			}
			else if (rows() == 1) {
				resize(1, n + 1);
				operator[](n) = x;
			}
			else {
				resize(n + 1, 1);
				operator[](n) = x;
			}

			return *this;
		}

	};

	using FP12 = FPX;

	// Fixed size array.
	template<int32_t N, int32_t M>
	struct fp12 {
		int32_t rows = N;
		int32_t columns = M;
		double array[N * M];
		fp12& reshape(int32_t r, int c)
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