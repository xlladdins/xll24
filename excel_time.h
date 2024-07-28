// excel_time.h - Excel Julian date to time_t conversion
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <expected>
#include <limits>
#include <timezoneapi.h>

namespace xll {

	// UTC = local time + bias
	inline std::expected<LONG, DWORD> timezone_bias()
	{
		DYNAMIC_TIME_ZONE_INFORMATION dtzi;
		DWORD ret = GetDynamicTimeZoneInformation(&dtzi);
		if (TIME_ZONE_ID_INVALID == ret) {
			return std::unexpected<DWORD>(ret);
		}

		return (dtzi.Bias + dtzi.DaylightBias) * 60; // in seconds
	}

	// Convert Excel Julian local date to UTC time time_t
	inline time_t to_time_t(double jd)
	{
		// Excel Julian date is days since 1900-01-00
		// time_t is seconds since 1970-01-01
		// 70 years, 17 leap years, 1 day
		const auto bias = timezone_bias();
		if (!timezone_bias()) {
			return -1;
		}

		return static_cast<time_t>((jd - 25569) * 86400 + *bias);
	}

	// Convert UTC time time_t to Excel local Julian date
	inline double from_time_t(time_t t)
	{
		// Excel Julian date is days since 1900-01-00
		// time_t is seconds since 1970-01-01
		// 70 years, 17 leap years, 1 day
		const auto bias = timezone_bias();
		if (!bias) {
			return std::numeric_limits<double>::quiet_NaN();
		}

		return static_cast<double>(t - *bias) / 86400 + 25569;
	}

} // namespace xll