// addin.h - Store data Excel needs to register add-ins.
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <algorithm>
#include <map>
#include "register.h"

namespace xll {

	// Create add-in to be registered with Excel.
	class AddIn {
		Args args;

		// Auto<Register> function with Excel
		void Register()
		{
			if (addins.contains(args.functionText)) {
				const auto err = OPER(L"AddIn: ")
					& args.functionText & OPER(L" already registered");
				XLL_WARNING(view(err));
			}
			else {
				addins[args.functionText] = &args;
				const Auto<xll::Register> xao_reg([&]() -> int {
					try {
						OPER regid = XlfRegister(&args);
						if (regid.xltype == xltypeNum) {
							regids[regid.val.num] = &args;
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
					addins.erase(text);
					regids.erase(Num(Excel(xlfEvaluate, text)));
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
		static inline std::map<OPER, Args*> addins;
		static inline std::map<double, Args*> regids;
		// Lookup using function text or register id.
		static inline const Args* find(const OPER& text)
		{
			const Args* pargs = nullptr;

			if (isNum(text)) {
				const auto i = regids.find(Num(text));
				if (i != regids.end()) {
					pargs = i->second;
				}
			}
			else {
				const auto i = addins.find(text);
				if (i != addins.end()) {
					pargs = i->second;
				}
			}

			return pargs;
		}
		
		AddIn(const Args& args_)
			: args(args_)
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
