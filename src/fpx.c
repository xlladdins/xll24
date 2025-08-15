// fpx.c - C VLA implementation of the fpx.h interface.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#include <stdlib.h>
#include <string.h>
#include "fpx.h"

void swap_double(double* x, double* y) {
	double t = *x;
	*x = *y;
	*y = t;
}

int32_t fpx_index(struct fpx* p, int32_t i, int32_t j)
{
	return p->columns * i + j;
}

struct fpx* fpx_malloc(int32_t r, int32_t c)
{
	struct fpx* fpx = malloc(sizeof(struct fpx) + (size_t)r * (size_t)c * sizeof(double));

	if (fpx) {
		fpx->rows = r;
		fpx->columns = c;
	}

	return fpx;
}

struct fpx* fpx_realloc(struct fpx* p, int32_t r, int32_t c)
{
	struct fpx* _p = realloc(p, sizeof(struct fpx) + (size_t)r * (size_t)c * sizeof(double));

	if (_p) {
		_p->rows = r;
		_p->columns = c;
	}

	return _p ? _p : (void*)0;
}

void fpx_free(struct fpx* p)
{
	free(p);
}

// TODO: use dimatcopy
struct fpx* fpx_transpose(struct fpx* fpx)
{
	int32_t r = fpx_rows(fpx);
	int32_t c = fpx_columns(fpx);
	if (r * c <= 1) {
		return fpx;
	}

	if (r > 1 && c > 1) {
		if (r == c) { // in-place transpose if square			
			for (int32_t i = 0; i < r; ++i) {
				for (int32_t j = i + 1; j < c; ++j) {
					swap_double(fpx->array + fpx_index(fpx, i, j), fpx->array + fpx_index(fpx, j, i));
				}
			}
		}
		else {
			struct fpx* _fpx = fpx_malloc(c, r);
			if (!_fpx) {
				return (struct fpx*)0;
			}
			for (int32_t i = 0; i < r; ++i) {
				for (int32_t j = i + 1; j < c; ++j) {
					swap_double(fpx->array + fpx_index(fpx, i, j), fpx->array + fpx_index(fpx, j, i));
				}
			}
			fpx_free(fpx);
			fpx = _fpx;
		}
	}
	fpx->rows = c;
	fpx->columns = r;

	return fpx;
}
