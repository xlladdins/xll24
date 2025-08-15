// fpx.h - Excel FP12 data type.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <stdint.h>

inline int32_t fpx_clamp(int32_t n, int32_t lo, int32_t hi)
{
	return n < lo ? lo : n > hi ? hi : n;
}

#pragma warning(push)
struct fpx {
	int32_t rows;
	int32_t columns;
	double array[1];
};
struct _fpx {
	int32_t rows;
	int32_t columns;
	double* array;
};

inline int32_t fpx_rows(struct fpx* fpx)
{
	return fpx->rows;
}
inline int32_t fpx_columns(struct fpx* fpx)
{
	return fpx->columns;
}	
inline int32_t fpx_size(struct fpx* fpx)
{
	return fpx->rows * fpx->columns;
}

// Row-major order
extern int32_t fpx_index(struct fpx* fpx, int32_t i, int32_t j);
struct fpx* fpx_malloc(int32_t r, int32_t c);
struct fpx* fpx_realloc(struct fpx* fpx, int32_t r, int32_t c);
void fpx_free(struct fpx*);
// in-place transpose
struct fpx* fpx_transpose(struct fpx* fpx);

