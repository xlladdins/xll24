// fpx.h - Excel FP12 data type.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once

#pragma warning(push)
struct fpx {
	int rows;
	int columns;
	double array[1];
};

inline int fpx_rows(struct fpx* fpx)
{
	return fpx->rows;
}
inline int fpx_columns(struct fpx* fpx)
{
	return fpx->columns;
}	
inline int fpx_size(struct fpx* fpx)
{
	return fpx->rows * fpx->columns;
}

// Row-major order
extern int fpx_index(struct fpx* fpx, int i, int j);
struct fpx* fpx_malloc(int r, int c);
struct fpx* fpx_realloc(struct fpx* fpx, int r, int c);
void fpx_free(struct fpx*);
// in-place transpose
struct fpx* fpx_transpose(struct fpx* fpx);

