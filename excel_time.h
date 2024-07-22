// excel_time.h - Excel Julian date to time_t conversion
#pragma once
#include <expected>
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
	constexpr time_t to_time_t(double jd)
	{
		// Excel Julian date is days since 1900-01-00
		// time_t is seconds since 1970-01-01
		// 70 years, 17 leap years, 1 day
		return static_cast<time_t>((jd - 25569) * 86400 + timezone_bias().value());
	}

	// Convert UTC time time_t to Excel local Julian date
	constexpr double from_time_t(time_t t)
	{
		// Excel Julian date is days since 1900-01-00
		// time_t is seconds since 1970-01-01
		// 70 years, 17 leap years, 1 day
		return static_cast<double>(t - timezone_bias().value()) / 86400 + 25569;
	}

} // namespace xll