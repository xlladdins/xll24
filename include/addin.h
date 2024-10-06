// addin.h - Store data Excel needs to register add-ins.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <algorithm>
#include <map>
#include "register.h"

namespace xll {

	inline std::map<double, Args*>& RegIds()
	{
		static std::map<double, Args*> regids;

		return regids;
	}

	// Create add-in to be registered with Excel.
	class AddIn {
		Args args;

		// Auto<Register> function with Excel
		void Register()
		{
			const Auto<xll::Register> xao_reg([&]() -> int {
				try {
					OPER regid = XlfRegister(&args);
					if (regid.xltype == xltypeNum) {
						RegIds()[regid.val.num] = &args;
					}
					else {
						const auto err = OPER(L"AddIn: failed to register: ") & args.functionText;
						XLL_WARNING(view(err));
					}

					return regid.xltype == xltypeNum;
				}
				catch (const std::exception& ex) {
					XLL_ERROR(ex.what());

					return false;
				}
				catch (...) {
					XLL_ERROR("AddIn::Auto<Register>: unknown exception");

					return false;
				}
			});
		}

		// Auto<Unregister> function with Excel
		void Unregister()
		{
			const Auto<xll::Unregister> xao_unreg([text = args.functionText]() {
				try {
					if (!XlfUnregister(text)) {
						const auto err = OPER(L"AddIn: failed to unregister: ") & text;
						XLL_WARNING(view(err));

						return FALSE;
					}
					RegIds().erase(Num(Excel(xlfEvaluate, text)));
				}
				catch (const std::exception& ex) {
					XLL_ERROR(ex.what());

					return FALSE;
				}
				catch (...) {
					XLL_ERROR("AddIn::Auto<Unregister>: unknown exception");

					return FALSE;
				}

				return TRUE;
			});
		}
	public:
		// Lookup using function text or register id.
		static Args* find(const OPER& text)
		{
			ensure(isStr(text) || isNum(text));

			Args* pargs = nullptr;
			double regid = 0;

			if (isStr(text)) {
				regid = Num(Excel(xlfEvaluate, text));
			}
			else {
				regid = Num(text);
			}
		
			const auto i = RegIds().find(regid);
			if (i != RegIds().end()) {
				pargs = i->second;
			}

			return pargs;
		}
		
		AddIn(const Args& args)
			: args(args)
		{
			Register();
			Unregister();
		}
		AddIn(const AddIn&) = delete;
		AddIn& operator=(const AddIn&) = delete;
		~AddIn()
		{ }
	};

} // namespace xll
