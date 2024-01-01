// fp.h - FP12 data type
#pragma once
#include <algorithm>
#include <initializer_list>
extern "C" {
#include "fpx.h"
}
#include "defines.h"

namespace xll {
	class FPX {
		struct fpx* _fpx;
	public:
		FPX()
			: _fpx(nullptr)
		{ }
		FPX(int r, int c)
			: _fpx(fpx_malloc(r, c))
		{
			ensure(_fpx->array);
		}
		FPX(int r, int c, const double* a)
			: FPX(r, c)
		{
			std::copy_n(a, r*c, _fpx->array);
		}
		FPX(std::initializer_list<double> a)
			: FPX(1, static_cast<int>(a.size()), a.begin())
		{ }
		FPX(const FPX&) = delete;
		FPX& operator=(const FPX&) = delete;
		~FPX()
		{
			fpx_free(_fpx);
		}

		// type punning
		operator const _FP12&() const
		{
			return reinterpret_cast<const _FP12&>(*_fpx);
		}

		constexpr int rows() const noexcept
		{
			return _fpx->rows;
		}
		constexpr int columns() const noexcept
		{
			return _fpx->columns;
		}
		constexpr int size() const noexcept
		{
			return rows() * columns();
		}

		FPX& resize(int r, int c)
		{
			ensure(nullptr != (_fpx = fpx_realloc(_fpx, r, c)));

			return *this;
		}
		constexpr FPX& reshape(int r, int c)
		{
			ensure(r * c == size());
			
			_fpx->rows = r;
			_fpx->columns = c;

			return *this;
		}

		double operator[](int i) const
		{
			return _fpx->array[i];
		}
		double& operator[](int i)
		{
			return _fpx->array[i];
		}

		double operator()(int i, int j) const
		{
			return _fpx->array[fpx_index(_fpx, i, j)];
		}
		double& operator()(int i, int j)
		{
			return _fpx->array[fpx_index(_fpx, i, j)];
		}
	};
}

constexpr bool operator==(const _FP12& a, const _FP12& b)
{
	return a.rows == b.rows 
		&& a.columns == b.columns 
		&& std::equal(a.array, a.array + a.rows*a.columns, b.array);
}