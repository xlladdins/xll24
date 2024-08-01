// fpx.c - C VLA implementation of the fpx.h interface.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#include <stdlib.h>
#include <string.h>
#include "fpx.h"

int fpx_index(struct fpx* p, int i, int j)
{
	return p->columns * i + j;
}

struct fpx* fpx_malloc(int r, int c)
{
	struct fpx* fpx = malloc(sizeof(struct fpx) + (size_t)r * (size_t)c * sizeof(double));

	if (fpx) {
		fpx->rows = r;
		fpx->columns = c;
	}

	return fpx;
}

struct fpx* fpx_realloc(struct fpx* p, int r, int c)
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
	int r = fpx_rows(fpx);
	int c = fpx_columns(fpx);
	int n = fpx_size(fpx);

	if (r > 1 && c > 1) {
		struct fpx* fpx_ = fpx_malloc(c, r);
		if (fpx_) {
			memcpy(fpx_->array, fpx->array, n * sizeof(double));
			for (int k = 1; k < n - 1; ++k) {
				// if 0 < k < n - 1 and k = c j + i then
				// r k % (n - 1) = (r c j + r i) % (n - 1) = j + (n - 1) j + r i % (n - 1) = j + r i
				fpx->array[(r * k) % (n - 1)] = fpx_->array[k];
			}
			fpx_free(fpx_);
		}
	}
	fpx->rows = c;
	fpx->columns = r;

	return fpx;
}
