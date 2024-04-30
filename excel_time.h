// excel_time.h - Excel Julian date to time_t conversion
#pragma once
#include <timezoneapi.h>

namespace xll {

	// Convert Excel Julian date to local time time_t
	constexpr time_t excel_time_t(double jd)
	{
		// Excel Julian date is days since 1900-01-00
		// time_t is seconds since 1970-01-01
		// 70 years, 17 leap years, 1 day
		return static_cast<time_t>((jd - 25569) * 86400);
	}

	// Convert local time time_t to Excel Julian date
	constexpr double time_t_excel(time_t t)
	{
		// Excel Julian date is days since 1900-01-00
		// time_t is seconds since 1970-01-01
		// 70 years, 17 leap years, 1 day
		return static_cast<double>(t) / 86400 + 25569;
	}

	// UTC = local time + bias
	inline auto timezone_bias()
	{
		static DYNAMIC_TIME_ZONE_INFORMATION dtzi;
		static DWORD init = GetDynamicTimeZoneInformation(&dtzi);

		return dtzi.Bias * 60; // in seconds
	}

} // namespace xll