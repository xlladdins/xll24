// fp.h - FP12 data type
#pragma once
#include <algorithm>
#include <initializer_list>
#include <span>
extern "C" {
#include "fpx.h"
}
#include "defines.h"

inline int rows(const FP12& a) noexcept
{
	return a.rows;
}
inline int columns(const FP12& a) noexcept
{
	return a.columns;
}
inline int size(const FP12& a) noexcept
{
	return a.rows * a.columns;
}

inline double* array(FP12& a) noexcept
{
	return a.array;
}
inline const double* array(const FP12& a) noexcept
{
	return a.array;
}

inline auto span(FP12& a) noexcept
{
	return std::span<double>(array(a), size(a));
}
inline auto span(const FP12& a) noexcept
{
	return std::span<const double>(array(a), size(a));
}

inline double* begin(FP12& a) noexcept
{
	return array(a);
}
inline const double* begin(const FP12& a) noexcept
{
	return array(a);
}
inline double* end(FP12& a) noexcept
{
	return array(a) + size(a);
}
inline const double* end(const FP12& a) noexcept
{
	return array(a) + size(a);
}

namespace xll {
	class FPX {
		struct fpx* _fpx;
	public:
		FPX() noexcept
			: _fpx(nullptr)
		{ }
		FPX(int r, int c)
			: _fpx(fpx_malloc(r, c))
		{
			ensure(_fpx);
		}
		// Copy from array having at least r*c elements.
		FPX(int r, int c, const double* a)
			: FPX(r, c)
		{
			std::copy_n(a, r*c, _fpx->array);
		}
		FPX(std::initializer_list<double> a)
			: FPX(1, static_cast<int>(a.size()), a.begin())
		{ }
		FPX(const FPX& a)
			: FPX(a.rows(), a.columns(), a.array())
		{ }
		FPX(FPX&& a) noexcept
			: _fpx(a._fpx)
		{
			a._fpx = nullptr;
		}
		FPX& operator=(const FPX& a)
		{
			if (this != &a) {
				fpx_free(_fpx);
				_fpx = fpx_malloc(a.rows(), a.columns());
				std::copy_n(a.array(), a.size(), _fpx->array);
			}

			return *this;
		}
		FPX& operator=(FPX&& a) noexcept
		{
			if (this != &a) {
				fpx_free(_fpx);
				_fpx = a._fpx;
				a._fpx = nullptr;
			}

			return *this;
		}
		~FPX()
		{
			fpx_free(_fpx);
		}

		explicit operator bool() const noexcept
		{
			return _fpx != nullptr;
		}

		// type punning
		operator const _FP12&() const noexcept
		{
			return reinterpret_cast<const _FP12&>(*_fpx);
		}
		const _FP12* get() const noexcept
		{
			return reinterpret_cast<const _FP12*>(_fpx);
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
		constexpr double* array() noexcept
		{
			return _fpx->array;
		}
		constexpr const double* array() const noexcept
		{
			return _fpx->array;
		}

		FPX& resize(int r, int c) noexcept
		{
			auto fpx_ = fpx_realloc(_fpx, r, c);
			if (fpx_) {
				_fpx = fpx_;
			}

			return *this;
		}
		constexpr FPX& reshape(int r, int c)
		{
			ensure(r * c == size());
			
			_fpx->rows = r;
			_fpx->columns = c;

			return *this;
		}

		double operator[](int i) const noexcept
		{
			return _fpx->array[i];
		}
		double& operator[](int i) noexcept
		{
			return _fpx->array[i];
		}

		double operator()(int i, int j) const noexcept
		{
			return _fpx->array[fpx_index(_fpx, i, j)];
		}
		double& operator()(int i, int j) noexcept
		{
			return _fpx->array[fpx_index(_fpx, i, j)];
		}

		std::span<double> span() noexcept
		{
			return std::span<double>(_fpx->array, size());
		}
		const std::span<double> span() const
		{
			return std::span<double>(_fpx->array, size());
		}

	};
}

inline const std::span<double> span(const xll::FPX& a)
{
	return a.span();
}

constexpr bool operator==(const _FP12& a, const _FP12& b)
{
	return a.rows == b.rows 
		&& a.columns == b.columns 
		&& std::equal(begin(a), end(a), begin(b), end(b));
}